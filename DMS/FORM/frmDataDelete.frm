VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form frmDataDelete 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Data obsolete System"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   Icon            =   "frmDataDelete.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   5415
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   2055
         _Version        =   65537
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   12632064
         BevelColorFace  =   8421376
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo2 
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   2055
         _Version        =   65537
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   12632064
         BevelColorFace  =   8421376
      End
      Begin VB.Label lblTo 
         BackColor       =   &H00C0B4A9&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblFrom 
         BackColor       =   &H00C0B4A9&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C000C0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdExecute 
      BackColor       =   &H0080C0FF&
      Caption         =   "Execute"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "frmDataDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsfactory             As ADODB.Recordset


Private Sub cmdExecute_Click()
'On Error GoTo ErrHandler
'     Dim idelete As Integer
'    idelete = MsgBox("Do you want to delete this record?", vbYesNo)
'    If idelete = vbYes Then
'            cn.Execute "Delete From Pat_Info_Main Where tmp_dt ='" & parseQuotes(txtSerial) & "'"
'            Call allClear
'    End If
'ErrHandler:
'    Select Case Err.Number
'        Case -2147217913
'            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
'     End Select
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

