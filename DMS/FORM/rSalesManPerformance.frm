VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form rSalesManPerformance 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diagnostic Management System"
   ClientHeight    =   1935
   ClientLeft      =   1785
   ClientTop       =   1740
   ClientWidth     =   6810
   Icon            =   "rSalesManPerformance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pre&view"
      Height          =   330
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   960
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Close"
      Height          =   330
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   930
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3930
      Top             =   270
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "2-Doctor Name"
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
   Begin VB.TextBox txtDoc_Name 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1050
      Width           =   4125
   End
   Begin VB.TextBox txtRefer_Code 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1230
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1050
      Width           =   1245
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3930
      Top             =   270
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSComCtl2.DTPicker stDt 
      Height          =   285
      Left            =   1230
      TabIndex        =   0
      Top             =   690
      Width           =   1240
      _ExtentX        =   2196
      _ExtentY        =   503
      _Version        =   393216
      Format          =   58851329
      CurrentDate     =   37306
   End
   Begin MSComCtl2.DTPicker edDt 
      Height          =   285
      Left            =   2820
      TabIndex        =   1
      Top             =   690
      Width           =   1240
      _ExtentX        =   2196
      _ExtentY        =   503
      _Version        =   393216
      Format          =   58851329
      CurrentDate     =   37337
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2520
      TabIndex        =   9
      Top             =   750
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Day between"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Man"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1050
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Man Performance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   270
      Width           =   2700
   End
End
Attribute VB_Name = "rSalesManPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdPreview_Click()
CRViewer1_MODE = 49
Viewer.Show vbModal
End Sub

Private Sub Form_Load()
'Adodc1.connectionstring = strcn.Connection
'Adodc1.RecordSource = "SELECT Emp_Name  From Emp_Info WHERE (Emp_ID = '" & txtRefer_Code & "') AND (Title = 'BD')"
'Adodc1.Refresh
'
'If Adodc1.Recordset.RecordCount > 0 Then
'   txtDoc_Name = Adodc1.Recordset!Doctor
'Else
'    txtDoc_Name = ""
'End If

stDt.value = Date
edDt.value = Date
End Sub

Private Sub txtRefer_Code_Change()
Adodc1.connectionstring = strcn.Connection
Adodc1.RecordSource = "SELECT Emp_Name  From Emp_Info WHERE (Emp_ID = '" & txtRefer_Code & "') AND (Title = 'BD')"
Adodc1.Refresh
    
If Adodc1.Recordset.RecordCount > 0 Then
   txtDoc_Name = Adodc1.Recordset!Emp_Name
Else
    txtDoc_Name = ""
End If
End Sub
