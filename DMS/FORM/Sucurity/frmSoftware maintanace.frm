VERSION 5.00
Begin VB.Form frmSoftware_Maintanance 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Software maintanace"
   ClientHeight    =   3960
   ClientLeft      =   2655
   ClientTop       =   1485
   ClientWidth     =   7215
   LinkTopic       =   "Form21"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSoftware maintanace.frx":0000
   ScaleHeight     =   3960
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   330
      Index           =   2
      Left            =   5925
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3375
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear"
      Height          =   330
      Index           =   1
      Left            =   5010
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3375
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      Height          =   330
      Index           =   0
      Left            =   4095
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3375
      Width           =   960
   End
   Begin VB.TextBox txtsoftware 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   3645
      TabIndex        =   7
      Top             =   1530
      Width           =   3255
   End
   Begin VB.TextBox txtscr_no 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   3645
      TabIndex        =   4
      Top             =   2070
      Width           =   3255
   End
   Begin VB.TextBox txtdescript 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   3645
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Software maintanance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   2295
      TabIndex        =   9
      Top             =   405
      Width           =   2790
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Software Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   2250
      TabIndex        =   8
      Top             =   1530
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   3
      Left            =   2250
      TabIndex        =   6
      Top             =   2115
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   2
      Left            =   2250
      TabIndex        =   5
      Top             =   2640
      Width           =   810
   End
End
Attribute VB_Name = "frmSoftware_Maintanance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
''Dim cn    As New ADODB.Connection
''Dim comm  As New ADODB.Command
'Private Sub Command1_Click(Index As Integer)
'Select Case Index
'    Case 0
'        If LTrim(RTrim(txtsoftware)) <> "" And LTrim(RTrim(txtscr_no)) <> "" Then
'            con.connectionstring = strcn.Connection
'            con.Open
'            cmd.CommandText = "exec add_scrn '" + ChkForQuote(txtsoftware) + "','" + ChkForQuote(txtscr_no) + "','" + ChkForQuote(txtdescript) + "'"
'            cmd.ActiveConnection = con
'            cmd.Execute
'            con.Close
'            txtscr_no = ""
'            txtdescript = ""
'        Else
'            MsgBox "Invalid entry", vbOKOnly, "Attention"
'            txtsoftware = ""
'            txtscr_no = ""
'            txtdescript = ""
'        End If
'        Case 1
'            txtsoftware = ""
'            txtscr_no = ""
'            txtdescript = ""
'        Case 2
'            Unload Me
'    End Select
'End Sub
'
