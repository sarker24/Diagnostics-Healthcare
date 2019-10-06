VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Object = "{EDF439C0-99E5-11CF-AFF3-004005100200}#1.2#0"; "pvmarq.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPatient_Info 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Memo Investigations [Unique Diagnostic Center]"
   ClientHeight    =   11145
   ClientLeft      =   -105
   ClientTop       =   435
   ClientWidth     =   15270
   FillColor       =   &H007DABD0&
   Icon            =   "Pat_Info.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc16 
      Height          =   375
      Left            =   12720
      Top             =   9960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "M Executive"
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
   Begin VB.TextBox txtDisbursment 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   9240
      Width           =   1455
   End
   Begin PVMarqueeLib.PVMarquee PVMarquee2 
      Height          =   495
      Left            =   0
      TabIndex        =   87
      Top             =   0
      Width           =   15375
      _Version        =   65538
      _ExtentX        =   27120
      _ExtentY        =   873
      _StockProps     =   29
      Text            =   "New Patient Information for OPD"
      ForeColor       =   65535
      BackColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Frame           =   5
      BeginProperty Font1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421376
      Text            =   "New Patient Information for OPD"
   End
   Begin VB.TextBox txtUName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0B4A9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   10200
      Width           =   1935
   End
   Begin VB.TextBox txtTotalAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0B4A9&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   84
      Top             =   10200
      Width           =   1695
   End
   Begin SSDataWidgets_A.SSDBCommand cmdSave 
      Height          =   495
      Left            =   4800
      TabIndex        =   14
      Top             =   10200
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   32768
      Font3D          =   1
      CaptionAlignment=   7
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Patient Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   240
      TabIndex        =   62
      Top             =   480
      Width           =   10455
      Begin VB.TextBox txtEmail 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1470
         TabIndex        =   8
         Top             =   3210
         Width           =   1230
      End
      Begin VB.TextBox txtMEName 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   2760
         TabIndex        =   90
         Top             =   3210
         Width           =   7500
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   2760
         TabIndex        =   19
         Top             =   2810
         Width           =   7500
      End
      Begin VB.TextBox txtCons_Code 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1470
         TabIndex        =   7
         Top             =   2810
         Width           =   1230
      End
      Begin VB.TextBox txtDoc_Name 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2000
         Width           =   7470
      End
      Begin VB.TextBox txtPat_Name 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2790
         TabIndex        =   0
         ToolTipText     =   "Write Patient's Name"
         Top             =   645
         Width           =   4185
      End
      Begin VB.TextBox txtAge 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4485
         MaxLength       =   17
         TabIndex        =   3
         Top             =   1140
         Width           =   2520
      End
      Begin VB.TextBox txtDoc_Addr 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1470
         MultiLine       =   -1  'True
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2400
         Width           =   8790
      End
      Begin VB.TextBox txtRefer_Code 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   6
         ToolTipText     =   "Doctor's ID"
         Top             =   2000
         Width           =   1230
      End
      Begin VB.TextBox txtPhone 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7725
         TabIndex        =   5
         Top             =   1140
         Width           =   2520
      End
      Begin VB.TextBox txtAddr 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1470
         TabIndex        =   4
         Top             =   1560
         Width           =   8790
      End
      Begin VB.ComboBox ComSex 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Pat_Info.frx":0CCA
         Left            =   1470
         List            =   "Pat_Info.frx":0CD4
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1140
         Width           =   1230
      End
      Begin VB.CommandButton Cr_DT_TM 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Current Date &Time"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7650
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   240
         Width           =   2550
      End
      Begin VB.TextBox txtDummy_Pat_ID 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   66
         TabStop         =   0   'False
         ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
         Top             =   270
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton cmdPatNew 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ne&w"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3090
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   270
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.CommandButton cmdPatOld 
         BackColor       =   &H00C0C0C0&
         Caption         =   "O&ld"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   270
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.CommandButton CmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   270
         Visible         =   0   'False
         Width           =   990
      End
      Begin MSComCtl2.DTPicker DT_TM 
         Height          =   285
         Left            =   8640
         TabIndex        =   68
         TabStop         =   0   'False
         ToolTipText     =   "Delevary Time"
         Top             =   630
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "HH:MM:SS"
         Format          =   51904514
         UpDown          =   -1  'True
         CurrentDate     =   37163
      End
      Begin MSComCtl2.DTPicker Dt 
         Height          =   285
         Left            =   7695
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   630
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   16777215
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   49152
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   51904515
         CurrentDate     =   37114
      End
      Begin VB.TextBox txtPat_ID1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
         Top             =   645
         Width           =   1350
      End
      Begin VB.TextBox txtPat_ID 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   69
         TabStop         =   0   'False
         ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
         Top             =   645
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lblMExecutive 
         BackColor       =   &H00C0B4A9&
         Caption         =   "M. Executive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   91
         Top             =   3210
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Consultant"
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
         TabIndex        =   82
         Top             =   2810
         Width           =   1215
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4110
         TabIndex        =   80
         Top             =   1140
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7110
         TabIndex        =   79
         Top             =   690
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference Code"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   78
         Top             =   2000
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   77
         Top             =   1140
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Patient  Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   76
         Top             =   630
         Width           =   1125
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7110
         TabIndex        =   75
         Top             =   1140
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   74
         Top             =   1590
         Width           =   615
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor's Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   73
         Top             =   2400
         Width           =   1305
      End
   End
   Begin VB.TextBox txtType 
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CheckBox Chkrefer_type 
      BackColor       =   &H00C0B4A9&
      Caption         =   "N&o"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   8280
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   "Select for different patient's Referance"
      Top             =   4080
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Test Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   240
      TabIndex        =   50
      Top             =   4200
      Width           =   10455
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   9000
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtS_Code 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   630
         MaxLength       =   3
         TabIndex        =   21
         Top             =   1080
         Width           =   705
      End
      Begin VB.TextBox nbrUnique_id 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   720
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox nbrTest_Rate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   5250
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1020
      End
      Begin VB.TextBox txtS_Name 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3720
      End
      Begin VB.TextBox txtM_Code 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         MaxLength       =   2
         TabIndex        =   20
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdDelete_TempTable 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9990
         MaskColor       =   &H007DABD0&
         Picture         =   "Pat_Info.frx":0CE6
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1380
         Width           =   390
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2745
         Left            =   120
         TabIndex        =   51
         Top             =   1320
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   4842
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         ColumnHeaders   =   0   'False
         ForeColor       =   16711680
         HeadLines       =   1
         RowHeight       =   19
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
            Size            =   9.75
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
               ColumnWidth     =   450.142
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3825.071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2505.26
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker Delv_TM 
         Height          =   255
         Left            =   7380
         TabIndex        =   23
         ToolTipText     =   "Delevary Time"
         Top             =   1080
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "HH:MM:SS"
         Format          =   51904514
         UpDown          =   -1  'True
         CurrentDate     =   37163
      End
      Begin MSComCtl2.DTPicker Delv_Dt 
         Height          =   255
         Left            =   6315
         TabIndex        =   18
         Top             =   1080
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   450
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   49152
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   51904515
         CurrentDate     =   37114
      End
      Begin MSForms.ComboBox cmbInvestigation 
         Height          =   450
         Left            =   120
         TabIndex        =   92
         Top             =   240
         Width           =   9735
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "17171;794"
         ColumnCount     =   5
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Arial"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   630
         TabIndex        =   81
         Top             =   870
         Width           =   705
      End
      Begin VB.Line Line2 
         X1              =   1440
         X2              =   1440
         Y1              =   1080
         Y2              =   1320
      End
      Begin VB.Line Line1 
         X1              =   600
         X2              =   600
         Y1              =   1080
         Y2              =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   59
         Top             =   870
         Width           =   420
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Test Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1440
         TabIndex        =   58
         Top             =   870
         Width           =   855
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   5250
         TabIndex        =   57
         Top             =   870
         Width           =   360
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   6315
         TabIndex        =   56
         Top             =   870
         Width           =   375
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7620
         TabIndex        =   55
         Top             =   870
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   1695
      Left            =   240
      TabIndex        =   27
      Top             =   8400
      Width           =   10455
      Begin VB.ComboBox txtFax 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Pat_Info.frx":15B0
         Left            =   6000
         List            =   "Pat_Info.frx":15BD
         TabIndex        =   12
         Text            =   "Doctor"
         Top             =   540
         Width           =   1455
      End
      Begin VB.TextBox nbrTotal 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox nbrDue 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8130
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   950
         Width           =   1020
      End
      Begin VB.TextBox nbrDisc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1155
         TabIndex        =   10
         Top             =   540
         Width           =   885
      End
      Begin VB.CheckBox ChkPaid 
         Caption         =   "Paid"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9210
         TabIndex        =   36
         Top             =   1275
         Width           =   1005
      End
      Begin VB.TextBox nbrDisc_Per 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3300
         TabIndex        =   11
         Top             =   540
         Width           =   930
      End
      Begin VB.TextBox nbrTot_Test 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   180
         Width           =   300
      End
      Begin VB.TextBox nbrNet_Amount 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1155
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   950
         Width           =   885
      End
      Begin VB.TextBox nbrAdv 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   950
         Width           =   1245
      End
      Begin VB.TextBox nbrVAT_Amt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   180
         Width           =   930
      End
      Begin VB.TextBox nbrVAT_Per 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1155
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   180
         Width           =   885
      End
      Begin VB.TextBox nbrTotal_Amt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9210
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   180
         Width           =   1020
      End
      Begin VB.TextBox nbrAdv_Pay 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3300
         TabIndex        =   13
         ToolTipText     =   "Advance Money"
         Top             =   950
         Width           =   930
      End
      Begin VB.TextBox nbrCollect_Fee 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9210
         TabIndex        =   24
         Top             =   540
         Width           =   1020
      End
      Begin VB.TextBox nbrTotCollect_Fee 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9210
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   950
         Width           =   1020
      End
      Begin VB.TextBox nbrTot_Disc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2265
         TabIndex        =   28
         Top             =   540
         Width           =   795
      End
      Begin VB.Label lblPaid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paid"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2640
         TabIndex        =   88
         Top             =   950
         Width           =   495
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Collection Fee"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   83
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4950
         TabIndex        =   49
         Top             =   180
         Width           =   420
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7710
         TabIndex        =   48
         Top             =   945
         Width           =   330
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   47
         Top             =   540
         Width           =   705
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4380
         TabIndex        =   46
         Top             =   540
         Width           =   165
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   45
         Top             =   950
         Width           =   960
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Advance"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4950
         TabIndex        =   44
         Top             =   950
         Width           =   1110
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT ( % )"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   43
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT Amount"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2100
         TabIndex        =   42
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total ( with VAT)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7680
         TabIndex        =   41
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount by"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4950
         TabIndex        =   40
         Top             =   540
         Width           =   930
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Fee"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   9210
         TabIndex        =   39
         Top             =   1020
         Width           =   1020
      End
   End
   Begin MSAdodcLib.Adodc Adodc15 
      Height          =   330
      Left            =   12720
      Top             =   9000
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Adodc15"
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
   Begin MSAdodcLib.Adodc Adodc14 
      Height          =   330
      Left            =   12720
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Adodc14"
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
   Begin MSAdodcLib.Adodc Adodc13 
      Height          =   330
      Left            =   12720
      Top             =   9000
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "Adodc13"
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
   Begin MSAdodcLib.Adodc Adodc12 
      Height          =   330
      Left            =   12720
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "12-DOCTOR ID"
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
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   12720
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Adodc11"
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
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   12720
      Top             =   8160
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3889
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
      Caption         =   "Adodc10"
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
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   12720
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "9-commission_main table"
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   12720
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "8-Unique_ID_select"
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   12720
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "7-show total ADVANCE"
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   12720
      Top             =   9000
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Caption         =   "6-show Discount+paid"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   12720
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "5-show advance"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   12720
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "4-M_CODE"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   12720
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "3-PatName"
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
      Left            =   12720
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "2-Doc Name"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   12720
      Top             =   8640
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "1-Ins+Upd-Pat_Info_Main"
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
   Begin SSDataWidgets_A.SSDBCommand cmdPrint 
      Height          =   495
      Left            =   5760
      TabIndex        =   15
      Top             =   10200
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   8421376
      Font3D          =   1
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand CmdPreview 
      Height          =   495
      Left            =   6720
      TabIndex        =   17
      Top             =   10200
      Width           =   1095
      _Version        =   196612
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Pre&view"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   8421376
      Font3D          =   1
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdNew 
      Height          =   495
      Left            =   7800
      TabIndex        =   16
      Top             =   10200
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&New"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   8421376
      Font3D          =   1
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdDelete 
      Height          =   495
      Left            =   8760
      TabIndex        =   25
      Top             =   10200
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   192
      Font3D          =   1
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdClose 
      Height          =   495
      Left            =   9720
      TabIndex        =   26
      Top             =   10200
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   8421376
      Font3D          =   1
      CaptionAlignment=   7
   End
   Begin PVMarqueeLib.PVMarquee PVMarquee1 
      Height          =   2055
      Left            =   10800
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   600
      Width           =   4335
      _Version        =   65538
      _ExtentX        =   7646
      _ExtentY        =   3625
      _StockProps     =   29
      Text            =   $"Pat_Info.frx":15D8
      ForeColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Direction       =   3
      Frame           =   5
      Justification   =   0
      BeginProperty Font1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12629161
      Text            =   $"Pat_Info.frx":1685
   End
End
Attribute VB_Name = "frmPatient_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Temp_Tab As New Recordset
Dim Temp_Table As New ADODB.Recordset
Dim Temp_Table_Helper As New ADODB.Recordset
Dim ChkPaidVal As String
Dim Total_Rate As Double 'for total test rate from temp_table
Dim Total_Test As Integer 'for total test from temp_table
Dim StrAdv_sum As String ' for show total Advance
Dim temp_open As String
Dim StrDATE As String
Dim StrTIME As String
Dim Date_TM As String 'for Add date and time
Dim CDate_TM As String
Dim CDate_TM2 As String ' date using only for pat_info_sub1_u
Dim CDate_TM3 As String 'date using only for Updpat_info_main
Dim CDate_TM4 As String 'date using only for Updpat_info_sub3
Dim CDate_TM5 As String 'date using only for Updpat_info_sub2
Dim CDate_TM6 As String 'date using only for Inspat_info_main
Dim CDate_TM7 As String 'date using only for Updpat_info_sub1
Dim CDate_TM8 As String 'date using only for Updpat_info_sub2
Dim CDate_TM9 As String 'date using only for Updpat_info_sub3
Dim CDate_TM10 As String
Dim StrRefer_Type As String 'for REFERENCE TYPE
Dim Del_Doc As String
Dim StPat_Type1 As Integer

Dim Strpat_id1 As String
Dim StrRow_Count As String
Dim StrPat_Type As String
Dim IntPat_ID As Double
Dim DblDisc As Double
Dim DummyPat_ID1 As String
Dim Strpat_MY As String


''----Add For Reporting Perpose----------------------------------------------


Private objReportApp                            As CRPEAuto.Application
Private objReport                               As CRPEAuto.Report
Private objReportDatabase                       As CRPEAuto.Database
Private objReportDatabaseTables                 As CRPEAuto.DatabaseTables
Private objReportDatabaseTable                  As CRPEAuto.DatabaseTable
Private objReportFormulaFieldDefinations        As CRPEAuto.FormulaFieldDefinitions
Private objReportFF                             As CRPEAuto.FormulaFieldDefinition


Private objReportSub                            As CRPEAuto.Report 'sub
Private objReportDatabaseSub                    As CRPEAuto.Database 'sub
Private objReportDatabaseTablesSub              As CRPEAuto.DatabaseTables 'sub
Private objReportDatabaseTableSub               As CRPEAuto.DatabaseTable 'sub
Private objReportFormulaFieldDefinationsSub     As CRPEAuto.FormulaFieldDefinitions
Private objReportFFSub                          As CRPEAuto.FormulaFieldDefinition


Private ObjPrinterSetting                       As CRPEAuto.PrintWindowOptions
 Private rscashmaster                           As ADODB.Recordset
'Private Tracer                              As Integer
Private strGroupName                            As String
Dim temp As Double
Dim temp1 As Double

Private rsTemp2                             As ADODB.Recordset

''--------------------------------------------------------------------------------


Private Sub ChkPaid_Click()
    
    If ChkPaid.value = 1 Then
        ChkPaidVal = "1"
    Else
    ChkPaidVal = "0"
    End If
End Sub

Private Sub Chkrefer_type_Click()
    Sel_Refer_Type
End Sub

Private Sub cmdAddItems_Click()
Frame5.Visible = True
PVTime1.Time = Time
DTPicker1.value = Date
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_GotFocus()
cmdClose.BevelColorFace = &HC0FFFF
End Sub

Private Sub cmdClose_LostFocus()
cmdClose.BevelColorFace = &H808000
End Sub

Private Sub cmdDelete_Click()
    Dim Strmsg As String
    Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
    
    If Strmsg = vbYes Then
    If txtPat_ID1.text <> "" Then
            If u_id <> "md" Then
                MsgBox "If you want to Delete the test, you should contact to Managing Director.....", vbCritical
            Exit Sub
            End If
    End If
    
    
    
        con.connectionstring = strcn.Connection
        con.Open
        Set cmd.ActiveConnection = con
        cmd.CommandText = "exec Pat_Info_SELECT 4,'" + txtPat_ID.text + "'"
        cmd.Execute
        con.Close

'        Temp_Table.Delete 'for delete from temporary table

        txtPat_Name = ""
        ComSex = "Male"
        txtRefer_Code = ""
        txtAddr = ""
        txtPhone = ""
        txtFax = ""
        txtEmail = ""
        Dt = Now

    End If
End Sub

Private Sub cmdDelete_GotFocus()
CmdDelete.BevelColorFace = &HC0FFFF
End Sub

Private Sub cmdDelete_LostFocus()
CmdDelete.BevelColorFace = &HC0&
End Sub

Private Sub cmdDelete_TempTable_Click()

    If txtPat_ID1.text <> "" Then
            If u_id <> "md" Then
                MsgBox "If you want to Delete the test, you should contact to Managing Director.....", vbCritical
            Exit Sub
            End If
    End If

    If Temp_Table.RecordCount <= 0 Then Exit Sub
    Dim Strmsg As String
    Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
    If Strmsg = vbYes Then
        DeletePat_Info_Sub1 'for DELETE from Pat_info_Sub1
        Temp_Table.Delete
 '++++++++++for count TOTAL_RATE from Temp_Table+++++++++
        Total_Rate = 0
        Temp_Table.MoveFirst
        While Temp_Table.EOF = False
                Total_Rate = Total_Rate + Temp_Table!test_rate
        Temp_Table.MoveNext
        Wend
        nbrTotal = Val(Total_Rate)
'++++++++++End count TOTAL_RATE from Temp_Table+++++++++
        
'=========count total test=============================
        Total_Test = 0
        Total_Test = Temp_Table.RecordCount
        Me.nbrTot_Test = Total_Test
'=========End count total test========================
    End If
        
End Sub

Private Sub cmdDelete_TempTable_GotFocus()
    cmdDelete_TempTable.BackColor = &HC0C0FF
End Sub

Private Sub cmdDelete_TempTable_LostFocus()
    cmdDelete_TempTable.BackColor = vbWhite
End Sub

Private Sub cmdNew_Click()
    
'txtUsrID.text = frmLogIn.Txtuserid.text
txtUName.text = u_id
    Temp_rst
    txtPat_ID1 = ""
    txtDummy_Pat_ID = ""
    txtPat_ID1.Visible = True
    txtPat_ID.Visible = False
    txtPat_ID = ""
    txtPat_Name = ""
    ComSex = "Male"
    txtAge = ""
    txtRefer_Code = ""
    txtAddr = ""
    txtPhone = ""
    txtFax = "Both"
    txtEmail = ""
    txtMEName = ""
    Dt.value = Now
    Delv_Dt.value = Now
    DT_TM.value = Now
    txtCons_Code = ""
    Text2 = ""
    txtDoc_Name = ""
    txtDegree = ""
    txtDoc_Addr = ""
    nbrAdv = 0
    nbrDisc = 0
    
    nbrTot_Disc = 0
    
    nbrDisc_Per = 0
    nbrDue = ""
    nbrNet_Amount = 0
    nbrTest_Rate = 0
    nbrTotal = ""
    nbrCollect_Fee = 0
    nbrTotCollect_Fee = 0
    nbrAdv_Pay.Locked = False
    nbrDisc.Locked = False
    nbrAdv_Pay = 0
    ChkPaid.value = 0
    nbrTot_Test = ""
    
    Chkrefer_type.value = 0
    
    nbrCollect_Fee.Locked = False
    cmdSearch.Visible = False
    cmdPatNew.Visible = False
    cmdPatOld.Visible = False
    
    Call Del_False_New_Doc
    
'    txtUsrID.text = frmLogIn.Txtuserid.text
    
    txtPat_Name.SetFocus
End Sub

Private Sub cmdNew_GotFocus()
cmdNew.BevelColorFace = &HC0FFFF
End Sub

Private Sub cmdNew_LostFocus()

cmdNew.BevelColorFace = &H808000

End Sub

Private Sub cmdPatNew_Click()

    Temp_rst
    txtCons_Code = ""
    Text2 = ""
    txtPat_Name = ""
    ComSex = "Male"
    txtAge = ""
    txtRefer_Code = ""
    txtAddr = ""
    txtPhone = ""
    txtFax = ""
    txtEmail = ""
    Dt.value = Now
    Delv_Dt.value = Now
    DT_TM.value = Now
    txtDoc_Name = ""
    txtDegree = ""
    txtDoc_Addr = ""
    nbrAdv = 0
    nbrDisc = 0
    
    nbrTot_Disc = 0
    
    nbrDisc_Per = 0
    nbrDue = ""
    nbrNet_Amount = 0
    nbrTest_Rate = 0
    nbrTotal = ""
    nbrCollect_Fee = 0
    nbrTotCollect_Fee = 0
    nbrAdv_Pay.Locked = False
    nbrDisc.Locked = False
    nbrAdv_Pay = 0
    ChkPaid.value = 0
    nbrTot_Test = ""
    
    Chkrefer_type.value = 0
    
    nbrCollect_Fee.Locked = False


    txtPat_ID1 = ""
    txtPat_ID = ""
    txtDummy_Pat_ID = ""
    txtPat_ID1.Visible = True
    txtPat_ID.Visible = False
    txtPat_ID1.SetFocus
    
End Sub

Private Sub cmdPatOld_Click()

    Temp_rst
    txtPat_Name = ""
    ComSex = "Male"
    txtAge = ""
    txtRefer_Code = ""
    txtAddr = ""
    txtPhone = ""
    txtFax = ""
    txtEmail = ""
    Dt.value = Now
    Delv_Dt.value = Now
    DT_TM.value = Now
    txtDoc_Name = ""
    txtDegree = ""
    txtDoc_Addr = ""
    txtCons_Code = ""
    Text2 = ""
    
    nbrAdv = 0
    nbrDisc = 0
    nbrTot_Disc = 0
    nbrDisc_Per = 0
    nbrDue = ""
    nbrNet_Amount = 0
    nbrTest_Rate = 0
    nbrTotal = ""
    nbrCollect_Fee = 0
    nbrTotCollect_Fee = 0
    nbrAdv_Pay.Locked = False
    nbrDisc.Locked = False
    nbrAdv_Pay = 0
    ChkPaid.value = 0
    nbrTot_Test = ""
    
    Chkrefer_type.value = 0
    
    nbrCollect_Fee.Locked = False



    txtPat_ID1 = ""
    txtPat_ID = ""
    txtDummy_Pat_ID = ""
    txtPat_ID1.Visible = False
    txtPat_ID.Visible = True
    txtPat_ID.SetFocus
End Sub

Private Sub CmdPreview_Click()
    
    Tracer = 0
    Call PrintReport


End Sub

Private Sub CmdPreview_GotFocus()
cmdPreview.BevelColorFace = &HC0FFFF
End Sub

Private Sub CmdPreview_LostFocus()
cmdPreview.BevelColorFace = &H808000
End Sub

Private Sub cmdPrint_Click()
    If StPat_ID = "" And txtPat_ID = "" Then Exit Sub
   
    Tracer = 1
Screen.MousePointer = vbHourglass
Call printReport1
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrint_GotFocus()
cmdPrint.BevelColorFace = &HC0FFFF
End Sub

Private Sub cmdPrint_LostFocus()
cmdPrint.BevelColorFace = &H808000
End Sub

Private Sub cmdQuit_Click()
Frame5.Visible = False
End Sub

Private Sub cmdSave_Click()
Strpat_id1 = "0"
    'MsgBox BoothN
    
    Dt.value = Now
    DT_TM.value = Now

    If Trim(txtPat_Name) = "" Then
        MsgBox "Paitent Name Mandatory"
        txtPat_Name.SetFocus
        Exit Sub
    End If
    
    If Trim(txtRefer_Code_Name) = "" Then
        MsgBox "Doctor's name Mandatory"
        txtRefer_Code = ""
        txtRefer_Code.SetFocus
        Exit Sub
    End If
    
    If Trim(txtPhone) = "" Then
        MsgBox "Entry Patient Phone No."
        txtPhone = ""
        txtPhone.SetFocus
        Exit Sub
    End If
   
    If Trim(nbrTotal_Amt) = "" Or Val(nbrTotal_Amt) = 0 Then
        MsgBox "Test Mandatory"
        txtM_Code.SetFocus
        Exit Sub
    End If
    
   
    'temp_rst 'RECORDSET
    Adodc1.connectionstring = strcn.Connection
    Adodc1.RecordSource = "select * from Pat_Info_main where pat_id='" & Trim(txtPat_ID.text) & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
    'MsgBox u_id
    
    If u_id <> "md" Then
        MsgBox "If you want to any change you should contact to Managing Director.., Your ID saved..", vbInformation
        Exit Sub
    End If
    
    
        
       StPat_ID = txtPat_ID 'TAKEN PAT_ID FOR PRINT
             
          
       Strpat_id1 = DummyPat_ID1
       
       If txtPat_ID1.text <> "" Then
             If StPat_Type1 <> Chkrefer_type.value Then
                 Make_Pat_ID1_U
'                 MsgBox Strpat_id1
            End If
                    
       End If
       
       
       'If Chkrefer_type = strRefer_Type1 Then
        '    Strpat_id1 = DummyPat_ID1
       'End If
       
       
    
       UpdPat_Info_Main
       Delete_Pat_Info_Sub1
       InsPat_Info_Sub1_U 'after delete data then INSERT
       InsPat_Info_Sub2_T1
       nbrAdv_Pay.Locked = False
       'UpdPat_Info_Sub3
       InsPat_Info_Sub3
       
       Search_Refer_Code 'search again refer_code for update refer_code/delete from doctor_info_new
       Del_New_Doc
       
    Else
    
        Make_Pat_ID1
        
        Dt.value = Now
        DT_TM.value = Now

        InsPat_Info_Main
    
    '///////SEARCH PATIENT ID for insert another table//////////////////////
        Adodc14.connectionstring = strcn.Connection
        Adodc14.RecordSource = "exec test_Info_SELECT 2,'" & BoothN & "'"
        Adodc14.Refresh
        If Adodc14.Recordset.RecordCount > 0 Then
        StPat_ID = Adodc14.Recordset!pat_id
        End If
    '///////END////////////////////////////////////////////
          
        InsPat_Info_Sub1
       ''''to insert into PAT_INFO_SUB2'''''''''
        If txtPat_ID = "" Then
            InsPat_Info_Sub2_T
            nbrAdv_Pay.Locked = False
        End If
    ''''''''''''''''end'''''''''''''''''''''''''
        InsPat_Info_Sub3
    ',,,,,,,,,for select,delete and insert into doctor_info_new,,,,,,,,,,,,,,,
        InsDoc_info_new
    ',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
    
    End If
    
    Temp_rst
    
    txtDummy_Pat_ID = ""
    txtPat_ID1.text = ""
'    txtPat_ID1.Visible = False
    txtPat_ID.text = ""
    txtPat_ID.Visible = False
    txtPat_Name = ""
    ComSex = "Male"
    txtAge = ""
    txtRefer_Code = ""
    txtCons_Code = ""
    Text2 = ""
    txtAddr = ""
    txtPhone = ""
    txtFax = ""
    txtEmail = ""
    txtMEName = ""
    Dt.value = Now
    'Delv_Dt.value = Now
    DT_TM.value = Now
    txtDoc_Name = ""
    txtDegree = ""
    txtDoc_Addr = ""
    nbrTot_Test = ""
    nbrTotal = ""
    nbrTotal_Amt = ""
    nbrDisc = 0
    nbrTot_Disc = 0
    nbrDisc_Per = 0
    nbrNet_Amount = 0
    nbrNet_Amount = ""
    nbrVAT_Amt = 0
    nbrTotal_Amt = ""
    nbrAdv.text = 0
    nbrAdv_Pay = 0
    nbrTotCollect_Fee.text = 0
    nbrCollect_Fee.text = 0
    nbrDue = ""
    ChkPaid.value = 0
    Chkrefer_type.value = 0
    '---------
    
    nbrCollect_Fee.Locked = False
    nbrDisc.Locked = False
    cmdPrint.SetFocus
   
End Sub
Private Sub InsPat_Info_Main()
    
    InsD_TM 'for insert current date and time
    Sel_Refer_Type 'for select REFERENCE TYPE
     
    
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_PAT_INFO_MAIN 'I','" + ChkForQuote(txtPat_Name.text) + "','" + ComSex.text + "','" + ChkForQuote(txtAge.text) + _
    "','" + txtRefer_Code.text + "','" + ChkForQuote(txtAddr.text) + "','" + txtPhone.text + _
    "','" + txtFax.text + "','" + txtEmail.text + "','" + u_id + _
    "','" + CDate_TM + _
    "'," + nbrVAT_Per.text + _
    "," + nbrVAT_Amt.text + _
    ",'" + BoothN + "','" + Format(Dt, "yyyy-mm-dd") + _
    "','" + CDate_TM + _
    "','" + StrRefer_Type + _
    "','" + Strpat_id1 + _
    "','" + Strpat_MY + "','" + txtCons_Code.text + "','" + ChkForQuote(Text2.text) + "'"
    
    cmd.Execute
'    MsgBox cmd.Execute
    con.Close
    
End Sub
Private Sub UpdPat_Info_Main()

      InsD_TM ' for insert current date and time
      
      Sel_Refer_Type 'for select REFERENCE TYPE

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_PAT_INFO_MAIN_UD 'U','" + txtPat_ID.text + _
    "','" + ChkForQuote(txtPat_Name) + "','" + ChkForQuote(ComSex) + "','" + ChkForQuote(txtAge) + "','" + txtRefer_Code + _
    "','" + ChkForQuote(txtAddr) + "','" + txtPhone + "','" + txtFax + _
    "','" + txtEmail + "','" + u_id + _
    "','" + CDate_TM + _
    "'," + nbrVAT_Per + "," + nbrVAT_Amt + ",'" + BoothN + _
    "','" + Format(CDate_TM3, "yyyy-mm-dd hh:mm") + _
    "','" + Format(CDate_TM6, "yyyy-mm-dd hh:mm") + _
    "','" + StrRefer_Type + _
    "','" + Strpat_id1 + _
    "','" + Strpat_MY + "','" + txtCons_Code.text + "','" + ChkForQuote(Text2) + "'"
    
    cmd.Execute
    con.Close
End Sub
Private Sub InsPat_Info_Sub1()

    
    Temp_Table.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
   
    While Temp_Table.EOF = False
          cmd.CommandText = "exec pro_PAT_INFO_SUB1 'I'," + StPat_ID + _
              ",'" + Temp_Table!m_code + _
              "','" + Temp_Table!s_code + _
              "'," + CStr(Temp_Table!test_rate) + _
              ",'" + Temp_Table!Delv_DTM + _
              "','" + Temp_Table!Type + _
              "','" + u_id + _
              "','" + CDate_TM + _
              "','" + Format(Dt, "yyyy-mm-dd") + _
              "','" + CDate_TM + _
              "','" + nbrUnique_id + "'"
'             Debug.Print cmd.CommandText
'             MsgBox cmd.CommandText
              cmd.Execute
              Temp_Table.MoveNext
              
    Wend
    con.Close

End Sub
Private Sub InsPat_Info_Sub1_U()


    If txtPat_ID = "" Then Exit Sub
    Temp_Table.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
   
    While Temp_Table.EOF = False
          cmd.CommandText = "exec pro_PAT_INFO_SUB1 'I'," + txtPat_ID + _
              ",'" + Temp_Table!m_code + _
              "','" + Temp_Table!s_code + _
              "'," + CStr(Temp_Table!test_rate) + _
              ",'" + Format(Temp_Table!Delv_DTM, "yyyy-mm-dd hh:mm") + _
              "','" + Temp_Table!Type + _
              "','" + u_id + _
              "','" + CDate_TM + _
              "','" + Format(CDate_TM2, "yyyy-mm-dd hh:mm") + _
              "','" + Format(CDate_TM7, "yyyy-mm-dd hh:mm") + _
              "','" + nbrUnique_id + "'"
'             Debug.Print cmd.CommandText
             'MsgBox cmd.CommandText
              cmd.Execute
              Temp_Table.MoveNext
    Wend
    con.Close

End Sub
Private Sub Delete_Pat_Info_Sub1()
    If txtPat_ID = "" Then Exit Sub
    
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec Pat_Info_Sub1_Delete 1,'" + Trim(txtPat_ID) + "'"
    cmd.Execute
    con.Close

End Sub
Private Sub DeletePat_Info_Sub1()

'    Temp_Table.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con

'    While Temp_Table.EOF = False
          cmd.CommandText = "exec Pat_Info_Sub1_Delete1 1,'" + Trim(nbrUnique_id) + "'"

              cmd.Execute
    con.Close
    txtM_Code = ""
    txtS_Code = ""
    txtS_Name = ""
    nbrRate = 0
    nbrUnique_id = ""
End Sub
Private Sub InsPat_Info_Sub3()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_PAT_INFO_SUB3 'I'," + StPat_ID + _
    "," + nbrDisc + "," + ChkPaidVal + ",'" + u_id + _
    "','" + CDate_TM + _
    "','" + Format(Dt, "yyyy-mm-dd") + _
    "','" + CDate_TM + "'"
    cmd.Execute
    con.Close
End Sub
Private Sub UpdPat_Info_Sub3()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_PAT_INFO_SUB3 'U'," + txtPat_ID.text + _
    "," + Round(nbrDisc) + "," + ChkPaidVal + ",'" + u_id + _
    "','" + CDate_TM + _
    "','" + Format(CDate_TM4, "yyyy-mm-dd") + _
    "','" + Format(CDate_TM9, "yyyy-mm-dd hh:mm") + "'"
    cmd.Execute
    con.Close
End Sub
Private Sub cmdSave_GotFocus()
    cmdSave.BevelColorFace = &HC0FFFF
End Sub
Private Sub cmdSave_LostFocus()
    cmdSave.BevelColorFace = &H8000&
End Sub

Private Sub cmdSearch_Click()

If u_id <> "md" Then Exit Sub
    
    
    Dim StrMMS As String
    StrMMS = MsgBox("Do you want Update New Patient ?", vbQuestion + vbYesNo)
    If StrMMS = vbYes Then
        cmdPatNew.Visible = True
        cmdPatOld.Visible = False
    Else
        cmdPatNew.Visible = False
        cmdPatOld.Visible = True
    End If
    
End Sub

Private Sub ComSex_GotFocus()
    ComSex.BackColor = &HFFFFC0
    
End Sub

Private Sub ComSex_LostFocus()
ComSex.BackColor = vbWhite
End Sub

Private Sub Cr_DT_TM_Click()
    Dt.value = Now
    DT_TM.value = Now
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
'    Sum_Rate
    nbrTot_Test = Rate_Tot
End Sub
Private Sub DataGrid1_DblClick()

    If Temp_Table.EOF = True Then Exit Sub
    
    txtM_Code = Temp_Table!m_code
    txtS_Code = Temp_Table!s_code
    txtS_Name = Temp_Table!s_name
    nbrTest_Rate = Temp_Table!test_rate
'    nbrUnique_id = Temp_Table_Helper!unique_id
    Select_Unique_ID
End Sub
Private Sub Delv_Dt_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If
    
End Sub

Private Sub Delv_TM_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If

End Sub

Private Sub Delv_TM_LostFocus()
    If Len(Trim(txtM_Code.text)) = 0 Then Exit Sub
    If Len(Trim(txtS_Code.text)) = 0 Then Exit Sub
    
'    Search_Type ' Search Type from table "Test_Info_Sub"
    
    '----------------check--------
    If Trim(nbrTest_Rate) = 0 Then
        MsgBox "Rate mandatory"
        nbrTest_Rate.SetFocus
        Exit Sub
    End If
        
    If Trim(txtS_Name.text) = "" Then
        MsgBox "Test Name mandatory"
        txtM_Code.text = ""
        txtM_Code.SetFocus
        Exit Sub
    End If
    
Dim Check As Integer
Check = 0
If Temp_Table.RecordCount > 0 Then
    Temp_Table.MoveFirst
    
        While Temp_Table.EOF = False
            If Temp_Table!m_code = txtM_Code And Temp_Table!s_code = txtS_Code Then
                Check = 1
            End If
    Temp_Table.MoveNext
        Wend
        
    If Check = 1 Then
        MsgBox "This Group Name and Test Name already exists"
        Check = 0
        txtS_Code.SetFocus
        Exit Sub
    End If
'    Temp_Table.MoveFirst
End If

'--------------end check-----

'++++++for insert Delivary Date and Time++++++++++++++

StrDATE = Trim(Format(Delv_Dt, "yyyy-mm-dd"))
StrTIME = Trim(Format(Delv_TM, "hh:mm"))

Date_TM = StrDATE + Space(1) + StrTIME
'MsgBox Date_TM
'++++++++++end+++++++++++++++++++++++++++++++++++++++
    
'+++to insert into TEMPORARY RECORDSET++
    
        Temp_Table.AddNew
        Temp_Table!m_code = txtM_Code
        Temp_Table!s_code = txtS_Code
        Temp_Table!s_name = txtS_Name
        Temp_Table!test_rate = nbrTest_Rate
        Temp_Table!Delv_DTM = Date_TM
        Temp_Table!Type = txtType
               
               
        'Search_Type ' Search Type from table "Test_Info_Sub"
        
        DataGrid1.Refresh
'++++++++++for count TOTAL_RATE from Temp_Table+++++++++
        Total_Rate = 0
        Temp_Table.MoveFirst
        While Temp_Table.EOF = False
                Total_Rate = Total_Rate + Temp_Table!test_rate
         
        Temp_Table.MoveNext
        Wend
        nbrTotal = Val(Total_Rate)
'++++++++++End count TOTAL_RATE from Temp_Table+++++++++
        
'=========count total test=============================
        Total_Test = 0
        Total_Test = Temp_Table.RecordCount
        Me.nbrTot_Test = Total_Test
'======================================================
    
        
'END ++++++++++++++++++++++++++++++++
        txtM_Code = ""
        txtS_Code = ""
        txtS_Name = ""
        nbrTest_Rate = 0
        txtType.text = ""
'       txtM_Code.SetFocus
       txtSearch.SetFocus
        
    DataGrid1.Columns(0).Width = 450.1418
    DataGrid1.Columns(1).Width = 810.1418
    DataGrid1.Columns(2).Width = 3825.071
    DataGrid1.Columns(3).Width = 1110.047
    DataGrid1.Columns(4).Width = 2100
    DataGrid1.Columns(5).Width = 800

    ChkPaid.value = 0
    
    nbrVAT_Amt = Val(nbrTotal) * Val(nbrVAT_Per) / 100 'for show VAT Amount
    
    
    
    
    
End Sub

Private Sub Form_DblClick()


'Call txtTotalAmount
    If cmdSearch.Visible = False Then
        cmdSearch.Visible = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    txtUName.text = u_id
    Temp_rst
    nbrAdv_Pay = 0
    nbrDisc = 0
    nbrTot_Disc.text = 0
    ChkPaidVal = 0
    nbrTotal = 0
    nbrTotCollect_Fee.text = 0
    nbrCollect_Fee.text = 0

    Delv_TM = Now
    
    txtUName.text = u_id
    
    cmbTestName
    ChkPaid.value = 0
    Dt.value = Now
    Delv_Dt.value = Now
    DT_TM.value = Now
    ComSex = "Male"
    temp_open = "0"
    Flush_VAT_Per
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Temp_Table = Nothing
End Sub

Private Sub cmbTestName()

Dim rsTemp2 As New ADODB.Recordset
con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con

rsTemp2.Open ("select distinct a.m_code,b.s_code,b.s_name,c.rate,b.type,b.SerialNo from " & _
               "test_info_main a,test_info_sub b, test_info_rate c " & _
               " WHERE a.m_code = b.m_code And a.m_code = c.m_code And b.s_code = c.s_code ORDER BY b.s_name ASC"), con, adOpenStatic, adLockReadOnly
                                     
    With cmbInvestigation
 .Clear

    While Not rsTemp2.EOF
       
       cmbInvestigation.ColumnCount = 3
       '~~> Set List Width
       cmbInvestigation.ListWidth = 500
      '~~> Set Column Widths/inch
      cmbInvestigation.ColumnWidths = "4 in; 2 in;2 in"
 
              cmbInvestigation.AddItem ""
              cmbInvestigation.List(cmbInvestigation.ListCount - 1, 0) = rsTemp2("s_name")
              cmbInvestigation.List(cmbInvestigation.ListCount - 1, 1) = rsTemp2("rate")
              cmbInvestigation.List(cmbInvestigation.ListCount - 1, 2) = rsTemp2("type")
                                
        rsTemp2.MoveNext
        
        Wend
    End With
    
    con.Close
End Sub


Private Sub cmbInvestigation_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 And cmbInvestigation.text = "" Then
SendKeys Chr(9)
End If

If KeyCode = 13 Then



    KeyCode = 0
    cmbInvestigation.SetFocus

    If Trim(cmbInvestigation) = "" Then Exit Sub

              Adodc16.connectionstring = strcn.Connection
              '    Adodc16.RecordSource = "exec  sp_found '" + txtM_Code + "','" + txtS_Code + "'"
              Adodc16.RecordSource = "select a.s_name,b.rate,a.type,a.m_code,a.s_code from test_info_sub a , test_info_rate b where a.m_code=b.m_code " & _
                                     "and a.s_code=b.s_code and a.s_name='" + cmbInvestigation.text + "'"
        Adodc16.Refresh

If Adodc16.Recordset.RecordCount > 0 Then
        txtS_Name = Adodc16.Recordset.Fields(0)
        nbrTest_Rate = Adodc16.Recordset.Fields(1)
        txtType.text = Adodc16.Recordset.Fields(2)
        txtM_Code.text = Adodc16.Recordset.Fields(3)
        txtS_Code.text = Adodc16.Recordset.Fields(4)

  End If
  End If

Call cmbInvestigationAdd
End Sub

Private Sub cmbInvestigationAdd()


    If Len(Trim(txtM_Code.text)) = 0 Then Exit Sub
    If Len(Trim(txtS_Code.text)) = 0 Then Exit Sub
    
'    Search_Type ' Search Type from table "Test_Info_Sub"
    
    '----------------check--------
    If Trim(nbrTest_Rate) = 0 Then
        MsgBox "Rate mandatory"
        nbrTest_Rate.SetFocus
        Exit Sub
    End If
        
    If Trim(txtS_Name.text) = "" Then
        MsgBox "Test Name mandatory"
        txtM_Code.text = ""
        txtM_Code.SetFocus
        Exit Sub
    End If
    
Dim Check As Integer
Check = 0
If Temp_Table.RecordCount > 0 Then
    Temp_Table.MoveFirst
    
        While Temp_Table.EOF = False
            If Temp_Table!m_code = txtM_Code And Temp_Table!s_code = txtS_Code Then
                Check = 1
            End If
    Temp_Table.MoveNext
        Wend
        
    If Check = 1 Then
        MsgBox "This Group Name and Test Name already exists"
        Check = 0
'        txtS_Code.SetFocus
        txtM_Code = ""
        txtS_Code = ""
        txtS_Name = ""
        nbrTest_Rate = 0
        txtType.text = ""
        cmbInvestigation.text = ""

        Exit Sub
    End If
'    Temp_Table.MoveFirst
End If

'--------------end check-----

'++++++for insert Delivary Date and Time++++++++++++++

StrDATE = Trim(Format(Delv_Dt, "yyyy-mm-dd"))
StrTIME = Trim(Format(Delv_TM, "hh:mm"))

Date_TM = StrDATE + Space(1) + StrTIME
'MsgBox Date_TM
'++++++++++end+++++++++++++++++++++++++++++++++++++++
    
'+++to insert into TEMPORARY RECORDSET++
    
        Temp_Table.AddNew
        Temp_Table!m_code = txtM_Code
        Temp_Table!s_code = txtS_Code
        Temp_Table!s_name = txtS_Name
        Temp_Table!test_rate = nbrTest_Rate
        Temp_Table!Delv_DTM = Date_TM
        Temp_Table!Type = txtType
               
               
        'Search_Type ' Search Type from table "Test_Info_Sub"
        
        DataGrid1.Refresh
'++++++++++for count TOTAL_RATE from Temp_Table+++++++++
        Total_Rate = 0
        Temp_Table.MoveFirst
        While Temp_Table.EOF = False
                Total_Rate = Total_Rate + Temp_Table!test_rate
         
        Temp_Table.MoveNext
        Wend
        nbrTotal = Val(Total_Rate)
'++++++++++End count TOTAL_RATE from Temp_Table+++++++++
        
'=========count total test=============================
        Total_Test = 0
        Total_Test = Temp_Table.RecordCount
        Me.nbrTot_Test = Total_Test
'======================================================
    
        
'END ++++++++++++++++++++++++++++++++
        txtM_Code = ""
        txtS_Code = ""
        txtS_Name = ""
        nbrTest_Rate = 0
        txtType.text = ""
        cmbInvestigation.text = ""
        
    DataGrid1.Columns(0).Width = 450.1418
    DataGrid1.Columns(1).Width = 810.1418
    DataGrid1.Columns(2).Width = 3825.071
    DataGrid1.Columns(3).Width = 1110.047
    DataGrid1.Columns(4).Width = 2100
    DataGrid1.Columns(5).Width = 800

    ChkPaid.value = 0
    
    nbrVAT_Amt = Val(nbrTotal) * Val(nbrVAT_Per) / 100 'for show VAT Amount
    
End Sub




Private Sub nbrAdv_Change()
'    nbrTot_Disc = Val(nbrTot_Disc) + Val(nbrDisc)

    nbrDue = Val(nbrNet_Amount) - Val(nbrAdv)
    '--for auto select PAID check box
    If Val(nbrNet_Amount) = 0 Then Exit Sub
    If Val(nbrAdv) = 0 Then Exit Sub
    If Val(nbrNet_Amount) = Val(nbrAdv) Then
        ChkPaid.value = 1
    Else
       If Val(nbrTotal_Amt) + Val(nbrTotCollect_Fee) = Val(nbrTot_Disc) Then
       ChkPaid.value = 1
       Else
       ChkPaid.value = 0
       End If
    End If
End Sub

Private Sub nbrAdv_GotFocus()

nbrAdv.BackColor = &HFFFFC0

End Sub

Private Sub nbrAdv_LostFocus()
    nbrAdv.BackColor = vbWhite
End Sub

Private Sub nbrAdv_Pay_Change()
    If Not IsNumeric(nbrAdv_Pay.text) Then
        MsgBox "Only Numaric value allow"
        nbrAdv_Pay = ""
        nbrAdv_Pay.SelStart = 0
        nbrAdv_Pay.SelLength = Len(nbrAdv_Pay)
        nbrAdv_Pay.SetFocus
    End If
End Sub

Private Sub nbrAdv_Pay_GotFocus()
    nbrAdv_Pay.BackColor = &HFFFFC0
    
    nbrAdv_Pay.SelStart = 0
    nbrAdv_Pay.SelLength = Len(nbrAdv_Pay)
End Sub

Private Sub nbrAdv_Pay_LostFocus()

    nbrAdv_Pay.BackColor = vbWhite
    
    If Trim(nbrAdv_Pay.text) = "" Or Val(nbrAdv_Pay) = 0 Then Exit Sub
        
        
    If Val(nbrAdv_Pay) > Val(nbrDue) Then
        MsgBox "It is Impossible to pay more then DUE", vbInformation
        nbrAdv_Pay.text = 0
        nbrAdv_Pay.SetFocus
        Exit Sub
    End If
    
    Dim Strmsg As String
    Strmsg = MsgBox("The Paitent's present DUE is  " + nbrDue + " and PAID " + nbrAdv + "  Do you want to pay more  " + nbrAdv_Pay + "", vbQuestion + vbYesNo)
        If Strmsg = vbYes Then
          ' If txtPat_ID = "" Then
'           InsPat_Info_Sub2
           nbrAdv_Pay.Locked = True
           nbrAdv = Val(nbrAdv) + Val(nbrAdv_Pay)
          ' End If
          cmdSave.SetFocus
        Else
            nbrAdv_Pay.text = "0"
            Exit Sub
        End If
        

End Sub

Private Sub nbrCollect_Fee_Change()
    If Not IsNumeric(nbrCollect_Fee.text) Then
        MsgBox "Only Numaric value allow"
        nbrCollect_Fee = 0
        nbrCollect_Fee.SelStart = 0
        nbrCollect_Fee.SelLength = Len(nbrCollect_Fee)
        nbrCollect_Fee.SetFocus
    End If

End Sub
Private Sub nbrCollect_GotFocus()
    nbrCollect_Fee.SelStart = 0
    nbrCollect_Fee.SelLength = Len(nbrCollect_Fee.text)
End Sub

Private Sub nbrCollect_Fee_GotFocus()
    nbrCollect_Fee.BackColor = &HFFFFC0
    
    nbrCollect_Fee.SelStart = 0
    nbrCollect_Fee.SelLength = Len(nbrCollect_Fee)
        
    'nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrDisc) + Val(nbrTotCollect_Fee.Text)
    nbrDisc.text = (Val(nbrDisc_Per) * Val(nbrTotal)) / 100 'for total discount
    nbrDisc = Round(Val(nbrDisc))
    'nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrDisc) + Val(nbrTotCollect_Fee.Text)
End Sub

Private Sub nbrCollect_Fee_LostFocus()

    nbrCollect_Fee.BackColor = vbWhite
    
    If Trim(nbrCollect_Fee.text) = "" Or Val(nbrCollect_Fee.text) = 0 Then Exit Sub
    
    Dim Strmsg As String
    Strmsg = MsgBox("The Paitent collection fee PAID : " + nbrTotCollect_Fee + "  Now he is going to pay :  " + nbrCollect_Fee + "", vbQuestion + vbYesNo)
    
        If Strmsg = vbYes Then
           nbrCollect_Fee.Locked = True
           nbrTotCollect_Fee = Val(nbrTotCollect_Fee) + Val(nbrCollect_Fee)
           'nbrCollect_Fee.Text = "0"
           nbrAdv_Pay.SetFocus
        Else
           nbrCollect_Fee.text = "0"
           Exit Sub
        End If
End Sub

Private Sub nbrDisc_Change()
    If Not IsNumeric(nbrDisc.text) Then
        MsgBox "Only Numaric value allow"
        nbrDisc = 0
        nbrDisc.SelStart = 0
        nbrDisc.SelLength = Len(nbrDisc)
        nbrDisc.SetFocus
    End If

    If Len(nbrTotal) = 0 Then Exit Sub
    
'    nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrDisc) + Val(nbrTotCollect_Fee.Text)

    If Val(nbrTotal) = 0 Then Exit Sub
    nbrDisc_Per.text = Val(nbrTot_Disc) * 100 / Val(nbrTotal) ' for percentence


    If Val(nbrTotal_Amt) + Val(nbrTotCollect_Fee) = Val(Me.nbrTot_Disc) Then
            ChkPaid.value = 1
        Else
            ChkPaid.value = 0
    End If
'
'    If Val(nbrDisc_Per.Text) = 0 Then
'       nbrDisc_Per.Text = ((Val(nbrDisc) * 100) / Val(nbrTotal.Text))
'    Else
'
'    End If

    nbrNet_Amount = Val(nbrTotal_Amt) - Round(Val(nbrDisc)) + Val(nbrTotCollect_Fee.text)
End Sub
Private Sub nbrDisc_GotFocus()

    nbrDisc.BackColor = &HFFFFC0
    
    nbrDisc.SelStart = 0
    nbrDisc.SelLength = Len(nbrDisc)
End Sub

Private Sub nbrDisc_LostFocus()
On Error Resume Next
nbrDisc.BackColor = vbWhite

If nbrDisc = "" Or nbrDisc = 0 Then Exit Sub
            
            Dim StrNbrDisc As String
            StrNbrDisc = Round(Val(nbrDisc.text))
           
           nbrTot_Disc = Val(nbrDisc)
           nbrDisc_Per.text = Round(Val(nbrDisc)) * 100 / Val(nbrTotal)
           nbrNet_Amount = Val(nbrTotal_Amt) - Round(Val(nbrDisc)) + Val(nbrTotCollect_Fee.text)
            
            nbrDisc.text = StrNbrDisc
'        Else
            
            nbrDisc.text = "0"

            Exit Sub
'        End If
End Sub

Private Sub nbrDisc_Per_Change()
    If Not IsNumeric(nbrDisc_Per.text) Then
        MsgBox "Only Numaric value allow"
        nbrDisc_Per = 0
        nbrDisc_Per.SelStart = 0
        nbrDisc_Per.SelLength = Len(nbrDisc_Per)
        nbrDisc_Per.SetFocus
    End If

    If Trim(nbrTotal) = 0 Then Exit Sub
    If Trim(nbrDisc) = 0 Then Exit Sub
    ' for percentence
    nbrDisc_Per.text = Round(Val(nbrDisc)) * 100 / Val(nbrTotal)
End Sub

Private Sub nbrDisc_Per_GotFocus()
    nbrDisc_Per.BackColor = &HFFFFC0
    
    
    nbrDisc_Per.SelStart = 0
    nbrDisc_Per.SelLength = Len(nbrDisc_Per)
End Sub

Private Sub nbrDisc_Per_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
    End If
End Sub

Private Sub nbrDisc_Per_LostFocus()

If Me.nbrDisc = "0" Then
    nbrDisc.text = (Val(nbrDisc_Per) * Val(nbrTotal)) / 100 'for total discount
    
'    nbrTot_Disc = Round(Val(nbrDisc.text)) + Val(nbrTot_Disc)
    nbrTot_Disc = Val(nbrDisc.text)
End If
    
    nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrTot_Disc) - Round(Val(nbrDisc)) + Val(nbrTotCollect_Fee.text)
    nbrDisc_Per.BackColor = vbWhite
End Sub

Private Sub nbrNet_Amount_Change()
    nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrTot_Disc) + Val(nbrTotCollect_Fee.text)
 
    nbrDue = Val(nbrNet_Amount) - Val(nbrAdv)
    If Val(nbrNet_Amount) = 0 Then Exit Sub
    If Val(nbrAdv) = 0 Then Exit Sub
    If Val(nbrNet_Amount) = Val(nbrAdv) Then
    ChkPaid.value = 1
    Else
        If Val(nbrTotal_Amt) + Val(nbrTotCollect_Fee) = Val(nbrTot_Disc) Then
            ChkPaid.value = 1
        Else
            ChkPaid.value = 0
        End If
    End If
    
End Sub

Private Sub nbrTest_Rate_Change()
    If Not IsNumeric(nbrTest_Rate.text) Then
        MsgBox "Only Numaric value allow"
        nbrTest_Rate = 0
        nbrTest_Rate.SelStart = 0
        nbrTest_Rate.SelLength = Len(nbrTest_Rate)
        nbrTest_Rate.SetFocus
    End If
End Sub

Private Sub nbrTest_Rate_GotFocus()
    nbrTest_Rate.BackColor = &HFFFFC0
    
    nbrTest_Rate.SelStart = 0
    nbrTest_Rate.SelLength = Len(nbrTest_Rate)
End Sub

Private Sub nbrTest_Rate_LostFocus()
    nbrTest_Rate.BackColor = vbWhite
End Sub

Private Sub nbrTot_Disc_Change()
nbrDisc.text = (Val(nbrDisc_Per) * Val(nbrTotal)) / 100 'for total discount
End Sub

Private Sub nbrTot_Disc_GotFocus()
nbrTot_Disc.BackColor = &HFFFFC0
End Sub

Private Sub nbrTot_Disc_LostFocus()
    Me.nbrTot_Disc.BackColor = vbWhite
End Sub

Private Sub nbrTotal_Amt_Change()


    nbrTotal_Amt = Val(nbrTotal) + Val(nbrVAT_Amt)
    nbrNet_Amount = Val(nbrTotal_Amt) - (Val(nbrTot_Disc) + Round(Val(nbrDisc))) + Val(nbrTotCollect_Fee.text)
    
    If Val(nbrTotal_Amt) + Val(nbrTotCollect_Fee) = Val(nbrTot_Disc) Then
            ChkPaid.value = 1
        Else
            ChkPaid.value = 0
    End If
    
    
End Sub
Private Sub nbrTotal_Change()

    nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrTot_Disc) + Val(nbrTotCollect_Fee.text)
    nbrVAT_Amt = Val(nbrTotal) * Val(nbrVAT_Per) / 100
    nbrTotal_Amt = Val(nbrTotal) + Val(nbrVAT_Amt)
End Sub

Private Sub nbrTotCollect_Fee_Change()

    nbrNet_Amount = Val(nbrTotal_Amt) - Val(nbrTot_Disc) + Val(nbrTotCollect_Fee.text)
    
    If Val(nbrTotal_Amt) + Val(nbrTotCollect_Fee) = Val(nbrTot_Disc) Then
            ChkPaid.value = 1
        Else
            ChkPaid.value = 0
    End If
    
End Sub

Private Sub nbrVAT_Amt_Change()

    nbrTotal_Amt = Val(nbrTotal) + Val(nbrVAT_Amt)
    nbrVAT_Amt = Round(nbrVAT_Amt, 0)

    
End Sub

Private Sub txtCons_Code_GotFocus()
txtRefer_Code.BackColor = &HFFFFC0
End Sub


Private Sub txtCons_Code_LostFocus()
On Error GoTo err_sub
    txtCons_Code.BackColor = vbWhite
    
    txtM_Code.TabStop = True
    If Trim(txtCons_Code) = "" Then Exit Sub
    'MsgBox "Patient1"
    Doc_List_MODE = "frmPatient_Info"

       If Trim(txtCons_Code.text) = "0" Then
       
            If Trim(txtCons_Code.text) <> "" Then
                
                If u_id <> "md" Then
                MsgBox "If you want to any change you should contact to Managing Director.., Your ID saved..", vbCritical
                txtRefer_Code = ""
                Exit Sub
                End If
                NdocMode = "0"
                frmDoctor_Info_New.txtPat_ID = txtPat_ID
            End If
            
            If Trim(txtPat_ID.text) = "" Then
                NdocMode = "1"
                frmDoctor_Info_New.txtPat_ID = "0"
            End If
       
       frmDoctor_Info_New.Show vbModal 'for new unknown doctor
       
       Else
               Adodc2.connectionstring = strcn.Connection
               
               Adodc2.RecordSource = "exec Pro_FLUSH2 1,'" & Trim(txtCons_Code.text) & "'"
               Adodc2.Refresh
               
               
                'MsgBox "Patient2"
               If Adodc2.Recordset.RecordCount > 0 Then
                   Text2.text = Adodc2.Recordset!doc_name
'                   txtDoc_Addr.Text = Adodc2.Recordset!addr
                   txtCons_Code.TabStop = True
               Else
               
                   'MsgBox "Patient3"
                   txtCons_Code.TabStop = False
                   frmDoc_List1.Show vbModal
                   Exit Sub
               End If
       End If
    Exit Sub
    
err_sub:
    MsgBox Err.Description

End Sub

Private Sub txtAddr_GotFocus()
txtAddr.BackColor = &HFFFFC0
End Sub

Private Sub txtAddr_LostFocus()

txtAddr.BackColor = vbWhite

End Sub

Private Sub txtAge_GotFocus()
txtAge.BackColor = &HFFFFC0
End Sub

Private Sub txtAge_LostFocus()
    txtAge.BackColor = vbWhite
End Sub

Private Sub txtDoc_Addr_GotFocus()
    txtDoc_Addr.BackColor = &HFFFFC0
End Sub

Private Sub txtDoc_Addr_LostFocus()
txtDoc_Addr.BackColor = vbWhite
End Sub

Private Sub txtDoc_Name_GotFocus()
    txtDoc_Name.BackColor = &HFFFFC0
End Sub

Private Sub txtDoc_Name_LostFocus()
txtDoc_Name.BackColor = vbWhite
End Sub

Private Sub txtEmail_GotFocus()
txtEmail.BackColor = &HFFFFC0
End Sub

Private Sub txtEmail_LostFocus()

    If Trim(txtEmail.text) = 0 Then Exit Sub
    
    Dim st As String
    Adodc16.connectionstring = strcn.Connection
    Adodc16.RecordSource = "exec Pro_FLUSH3 1,'" & Trim(txtEmail.text) & "'"
    Adodc16.Refresh
    
    
    If Adodc16.Recordset.RecordCount > 0 Then
      txtMEName.text = Adodc16.Recordset!Emp_Name
        
        End If
    Exit Sub

End Sub

Private Sub txtFax_GotFocus()

txtFax.BackColor = &HFFFFC0

End Sub

Private Sub txtFax_LostFocus()
txtFax.BackColor = vbWhite
End Sub

Private Sub txtM_Code_GotFocus()
txtM_Code.BackColor = &HFFFFC0
End Sub

Private Sub txtM_Code_LostFocus()

    On Error GoTo err_sub
    
    txtM_Code.BackColor = vbWhite
    
    Test_List_Mode = "frmPatient_Info_M" 'mode for show 'TEST NAME LIST'
    
   
    If Trim(txtM_Code.text) = "" Then
        If Val(nbrTotal) <> 0 Then
            nbrDisc.SetFocus
        End If
        Exit Sub
    End If
    
    Adodc4.connectionstring = strcn.Connection
    Adodc4.RecordSource = "exec  sp_found '" + txtM_Code + "',''"
    Adodc4.Refresh

    If Adodc4.Recordset.Fields(0) = "N" Then
     frmTest_List.Show vbModal 'show TEST NAME order by s_code
     Exit Sub
    End If
    Exit Sub
    
err_sub:
    MsgBox Err.Description
End Sub

Private Sub txtPat_ID_Change()
    If Trim(txtPat_ID) = "" Then Exit Sub
    If Not IsNumeric(txtPat_ID.text) Then
        MsgBox "Invalid Patient ID, Please try again.......  "
        txtPat_ID = ""
        txtPat_ID.SelStart = 0
        txtPat_ID.SelLength = Len(txtPat_ID)
        txtPat_ID.SetFocus
    End If

End Sub
Private Sub txtPat_ID_GotFocus()
    txtPat_ID.SelStart = 0
    txtPat_ID.SelLength = Len(txtPat_ID)
End Sub

Private Sub txtPat_ID_LostFocus()
'
    
    ChkPaid.value = 0
    Temp_rst
    StrAdv_sum = 0
    nbrAdv.text = ""
   '-----------------------------------------------------------
    DataGrid1.Columns(0).Width = 450.1418
    DataGrid1.Columns(1).Width = 810.1418
    DataGrid1.Columns(2).Width = 3825.071
    DataGrid1.Columns(3).Width = 1110.047
    DataGrid1.Columns(4).Width = 2100
    DataGrid1.Columns(5).Width = 800
    '-----------------------------------------------------------
        
    If Len(Trim(txtPat_ID.text)) = 0 Then Exit Sub

       Adodc3.connectionstring = strcn.Connection
       Adodc3.RecordSource = "exec Pat_Info_SELECT 1," + txtPat_ID + ""
       Adodc3.Refresh
       If Adodc3.Recordset.RecordCount > 0 Then
          txtPat_ID.text = Adodc3.Recordset!pat_id
          txtPat_Name = Adodc3.Recordset!pat_name
          ComSex = Adodc3.Recordset!Sex
          txtAge = Adodc3.Recordset!age

          txtAddr = Adodc3.Recordset!addr
          txtPhone = Adodc3.Recordset!phone
          txtFax = Adodc3.Recordset!fax
          txtEmail = Adodc3.Recordset!email

          nbrVAT_Per = Adodc3.Recordset!vat_per
          nbrVAT_Amt = Adodc3.Recordset!vat_amt

            '`````````````to show date and time from pat_info_main``````
           Adodc11.connectionstring = strcn.Connection
           Adodc11.RecordSource = "exec Pat_Info_SELECT 1,'" + txtPat_ID + "'"
           Adodc11.Refresh

            Dim StrCdt1 As String
            Dim StrCtime1 As String
            Dim CDate_TM1 As String

           If Adodc11.Recordset.RecordCount > 0 Then
            CDate_TM1 = Adodc11.Recordset!Dt
            CDate_TM3 = Adodc11.Recordset!tmp_dt
            CDate_TM6 = Adodc11.Recordset!dt1
            
            StrCdt1 = Mid(CDate_TM1, 1, 10)
            StrCtime1 = Mid(CDate_TM1, 12, 12)
            Dt = StrCdt1
            DT_TM = StrCtime1
'
            End If
            
     '```````END````````````````````````````````````````````````
            
     '`````````````to show date and time from pat_info_sub1``````
           Adodc11.connectionstring = strcn.Connection
           Adodc11.RecordSource = "exec Pat_Info_SELECT 5,'" + txtPat_ID + "'"
           Adodc11.Refresh

           If Adodc11.Recordset.RecordCount > 0 Then
            CDate_TM2 = Adodc11.Recordset!tmp_dt
            CDate_TM7 = Adodc11.Recordset!dt1
           End If
      '`````````````````END```````````````````````````
      
      '`````````````to show date and time from pat_info_sub2``````
           Adodc11.connectionstring = strcn.Connection
           Adodc11.RecordSource = "exec Pat_Info_SELECT 2,'" + txtPat_ID + "'"
           Adodc11.Refresh

           If Adodc11.Recordset.RecordCount > 0 Then
            CDate_TM5 = Adodc11.Recordset!tmp_dt
            CDate_TM8 = Adodc11.Recordset!dt1
           End If
      '`````````````````END```````````````````````````
      
      '`````````````to show date and time from pat_info_sub3``````
           Adodc11.connectionstring = strcn.Connection
           Adodc11.RecordSource = "exec Pat_Info_SELECT 3,'" & txtPat_ID & "'"
           Adodc11.Refresh
          If Adodc11.Recordset.RecordCount > 0 Then
            CDate_TM4 = Adodc11.Recordset!tmp_dt
            CDate_TM9 = Adodc11.Recordset!dt1
           End If
      '`````````````````END```````````````````````````
               
        
        
           '--------flush into Temp_Tabel-------------------------------
            con.connectionstring = strcn.Connection
            con.Open
            
            Temp_Table_Helper.Open "select m_code,s_code,(select s_name from test_info_sub Where test_info_sub.s_code = pat_info_sub1.s_code and test_info_sub.m_code=pat_info_sub1.m_code and pat_id='" + txtPat_ID + "') as s_name,test_rate,delv_dt,type,unique_id from pat_info_sub1 where pat_id='" + txtPat_ID + "'", con
            
            'MsgBox Temp_Table_Helper.RecordCount
              While Temp_Table_Helper.EOF = False
                    Temp_Table.AddNew
                                                            
                    Temp_Table!m_code = Temp_Table_Helper!m_code
                    Temp_Table!s_code = Temp_Table_Helper!s_code
                    Temp_Table!s_name = Temp_Table_Helper!s_name
                    Temp_Table!test_rate = Temp_Table_Helper!test_rate
                    Temp_Table!Delv_DTM = Temp_Table_Helper!Delv_Dt
                    Temp_Table!Type = Temp_Table_Helper!Type
                    Temp_Table_Helper.MoveNext
              Wend
                
            DataGrid1.Refresh
            Temp_Table_Helper.Close
            con.Close
           
           
           '---------------------------------------------------------
               '`````````````to show DISCOUNT from pat_info_sub3``````
               Adodc6.connectionstring = strcn.Connection
               Adodc6.RecordSource = "exec Pat_Info_SELECT 11,'" & txtPat_ID.text & "'"
               Adodc6.Refresh
    
               If Adodc6.Recordset.RecordCount > 0 Then
               Dim strchkpaid As String
                nbrDisc.text = "0"
                
                nbrTot_Disc = Adodc6.Recordset!disc
                strchkpaid = Adodc6.Recordset!paid
                'MsgBox strchkpaid
                    If Trim(strchkpaid) = "True" Then
                    ChkPaid.value = 1
                    ChkPaidVal = "1"
                    Else
                    ChkPaid.value = 0
                    ChkPaidVal = "0"
                    End If
               End If
               '```````````````````````````````````````````````````````
               
               '`````````````to show REFERENCE_TYPE from pat_info_MAIN``````
               Adodc6.connectionstring = strcn.Connection
               Adodc6.RecordSource = "exec Pat_Info_SELECT 1,'" + txtPat_ID + "'"
               Adodc6.Refresh
    
               If Adodc6.Recordset.RecordCount > 0 Then
               Dim strRefer_Type1 As String
               
                strRefer_Type1 = Adodc6.Recordset!refer_type
                    If strRefer_Type1 = 1 Then
                    Chkrefer_type.value = 1
                    strRefer_Type1 = "1"
                    Else
                    Chkrefer_type.value = 0
                    strRefer_Type1 = "0"
                    End If
               End If
               '``````````````````````````````````````````````````````
               
               '*************for flush doctor ID and name ****************
               Adodc12.connectionstring = strcn.Connection
               Adodc12.RecordSource = "exec Pat_Info_SELECT 7,'" + txtPat_ID + "'"
               
               Adodc12.Refresh
               If Adodc12.Recordset.RecordCount > 0 Then
               
                   txtRefer_Code = Adodc12.Recordset!Refer_code
                'MsgBox txtRefer_Code
               
               End If
               
               
'              '======DONTOR NAME FROM DOCTOR_INFO_NEW=============
               Adodc13.connectionstring = strcn.Connection
               Adodc13.RecordSource = "exec Pat_Info_SELECT 6,'" + txtPat_ID + "'"
               
               Adodc13.Refresh
               If Adodc13.Recordset.RecordCount > 0 Then
               
                  txtDoc_Name = Adodc13.Recordset!doc_name
                  txtDoc_Addr = Adodc13.Recordset!addr
               End If
               '=====================END===========================
               ',,,,,,,,,,,,,,for get registered doctor,,,,,,,,,,,
               Dim My_Rst As New ADODB.Recordset
               con.connectionstring = strcn.Connection
               con.Open
               Set My_Rst.ActiveConnection = con
               My_Rst.Open "exec Pro_FLUSH1 1,'" & Trim(txtRefer_Code.text) & "'", con
               If My_Rst.EOF = False Then
               
                    txtDoc_Name.text = My_Rst!doc_name
                    txtDoc_Addr.text = My_Rst!addr
               Else
                    txtDoc_Name.ForeColor = vbBlack
                    txtDoc_Addr.ForeColor = vbBlack
               End If
               My_Rst.Close
               con.Close
               
                
               ',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
              
               '***************end****************************************
                
         Else
           txtPat_Name = ""
           ComSex = "Male"
           txtAge = ""
           txtRefer_Code = ""
           txtDegree = ""
           txtAddr = ""
           txtPhone = ""
           txtFax = ""
           txtEmail = ""
           Dt.value = Now
           Delv_Dt.value = Now
           nbrVAT_Amt = 0
           nbrAdv = 0
           nbrDisc = 0
           nbrTot_Disc = 0
           nbrDisc_Per = 0
           nbrDue = ""
           nbrNet_Amount = ""
           
           nbrTest_Rate = ""
           nbrTotal = ""
           ChkPaid.value = 0
           Delv_TM.value = Now
           Chkrefer_type.value = 0
        End If
        
'++++++++++for count TOTAL_RATE from Temp_Table+++++++++
        If Temp_Table.RecordCount > 0 Then
        Total_Rate = 0
        Temp_Table.MoveFirst
        While Temp_Table.EOF = False
                Total_Rate = Total_Rate + Temp_Table!test_rate

        Temp_Table.MoveNext
        Wend
        nbrTotal = Total_Rate
        End If
'++++++++++End count TOTAL_RATE from Temp_Table+++++++++
        
'=========count total test=============================
        Total_Test = 0
        Total_Test = Temp_Table.RecordCount
        Me.nbrTot_Test = Total_Test
'=========End count total test========================

'>>>>>>>>>>>>>>>>to show total advance>>>>>>>>>>>>>>>>>>>>>>
    Adodc7.connectionstring = strcn.Connection
    Adodc7.RecordSource = "exec Pro_FLUSH 3,'" & Trim(txtPat_ID.text) & "'"
    Adodc7.Refresh
    If Adodc7.Recordset.RecordCount > 0 Then
        nbrAdv.text = Adodc7.Recordset!adv_sum
        nbrTotCollect_Fee.text = Adodc7.Recordset!Coll_sum
    End If
'<<<<<<<<<<<<End show total advance<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'nbrDisc_Per.Text = Val(nbrDisc) * 100 / Val(nbrTotal) ' for percentence

    DataGrid1.Columns(0).Width = 450.1418
    DataGrid1.Columns(1).Width = 810.1418
    DataGrid1.Columns(2).Width = 3825.071
    DataGrid1.Columns(3).Width = 1110.047
    DataGrid1.Columns(4).Width = 2100
    DataGrid1.Columns(5).Width = 800

nbrAdv_Pay.SetFocus
End Sub

Private Sub txtPat_ID1_LostFocus()
On Error Resume Next

    If txtPat_ID1 = "" Then Exit Sub
    If txtPat_ID1 <> "" Then
        txtPat_ID.TabStop = False
    End If
    Search_Patient_Type
    
    If StrRow_Count > "1" Then
        
            Dim Patmsg As String
            Patmsg = MsgBox("Do you want to update Inside Patient's information ? ", vbQuestion + vbYesNo)
            If Patmsg = vbYes Then
                StrPat_Type = "0"
                
                Srch_Pat_ID
            Else
                StrPat_Type = "1"
                
                Srch_Pat_ID
                
            End If
    Else
            Srch_Pat_ID1
    End If
    
   
    txtPat_ID = IntPat_ID
    
    txtDummy_Pat_ID.text = IntPat_ID
    
    If IntPat_ID = 0 Then
        MsgBox "Invalid ID, Try again"
        txtPat_ID = ""
        txtPat_ID1 = ""
        txtDummy_Pat_ID = ""
        txtPat_ID1.SetFocus
        Exit Sub
    End If
    
    
    '---------------------------------------------------------------
    
    ChkPaid.value = 0
    Temp_rst
    StrAdv_sum = 0
    nbrAdv.text = ""
   '-----------------------------------------------------------
    DataGrid1.Columns(0).Width = 450.1418
    DataGrid1.Columns(1).Width = 810.1418
    DataGrid1.Columns(2).Width = 3825.071
    DataGrid1.Columns(3).Width = 1110.047
    DataGrid1.Columns(4).Width = 2100
    DataGrid1.Columns(5).Width = 800
    '-----------------------------------------------------------
        
    If Len(Trim(txtPat_ID.text)) = 0 Then Exit Sub
      'for flush patient information
       Adodc3.connectionstring = strcn.Connection
       Adodc3.RecordSource = "exec Pat_Info_SELECT 1," & txtDummy_Pat_ID.text & ""
       Adodc3.Refresh
       If Adodc3.Recordset.RecordCount > 0 Then
          txtPat_ID.text = Adodc3.Recordset!pat_id
          txtPat_Name = Adodc3.Recordset!pat_name
          ComSex = Adodc3.Recordset!Sex
          txtAge = Adodc3.Recordset!age
          txtAddr = Adodc3.Recordset!addr
          txtPhone = Adodc3.Recordset!phone
          txtFax = Adodc3.Recordset!fax
          txtEmail = Adodc3.Recordset!email
          txtUName = Adodc3.Recordset!uid
          nbrVAT_Per = Adodc3.Recordset!vat_per
          nbrVAT_Amt = Adodc3.Recordset!vat_amt
          StPat_Type1 = Adodc3.Recordset!refer_type
          DummyPat_ID1 = Adodc3.Recordset!pat_id1
          Strpat_MY = Adodc3.Recordset!pat_my
          txtCons_Code = Adodc3.Recordset!Cons_Code
          Text2 = Adodc3.Recordset!cons
          txtUsrID = Adodc3.Recordset!uid
          
'          MsgBox DummyPat_ID1
'          MsgBox Strpat_MY
          
            '`````````````to show date and time from pat_info_main``````
           Adodc11.connectionstring = strcn.Connection
           Adodc11.RecordSource = "exec Pat_Info_SELECT 1,'" + txtDummy_Pat_ID.text + "'"
           Adodc11.Refresh

            Dim StrCdt1 As String
            Dim StrCtime1 As String
            Dim CDate_TM1 As String

           If Adodc11.Recordset.RecordCount > 0 Then
            CDate_TM1 = Adodc11.Recordset!Dt
            CDate_TM3 = Adodc11.Recordset!tmp_dt
            CDate_TM6 = Adodc11.Recordset!dt1
            
            StrCdt1 = Mid(CDate_TM1, 1, 10)
            StrCtime1 = Mid(CDate_TM1, 12, 12)
            Dt = StrCdt1
            DT_TM = StrCtime1
'
            End If
            
           '```````END````````````````````````````````````````````````
            
     '`````````````to show date and time from pat_info_sub1``````
           Adodc11.connectionstring = strcn.Connection
           Adodc11.RecordSource = "exec Pat_Info_SELECT 5,'" + txtDummy_Pat_ID.text + "'"
           Adodc11.Refresh

           If Adodc11.Recordset.RecordCount > 0 Then
            CDate_TM2 = Adodc11.Recordset!tmp_dt
            CDate_TM7 = Adodc11.Recordset!dt1
           End If
      '`````````````````END```````````````````````````
      
      '`````````````to show date and time from pat_info_sub2``````
           Adodc11.connectionstring = strcn.Connection
           Adodc11.RecordSource = "exec Pat_Info_SELECT 2,'" + txtDummy_Pat_ID.text + "'"
           Adodc11.Refresh

           If Adodc11.Recordset.RecordCount > 0 Then
            CDate_TM5 = Adodc11.Recordset!tmp_dt
            CDate_TM8 = Adodc11.Recordset!dt1
           End If
      '`````````````````END```````````````````````````
      
      '`````````````to show date and time from pat_info_sub3``````
           Adodc11.connectionstring = strcn.Connection
           Adodc11.RecordSource = "exec Pat_Info_SELECT 3,'" & txtDummy_Pat_ID.text & "'"
           Adodc11.Refresh
          If Adodc11.Recordset.RecordCount > 0 Then
            CDate_TM4 = Adodc11.Recordset!tmp_dt
            CDate_TM9 = Adodc11.Recordset!dt1
           End If
      '`````````````````END```````````````````````````
               
        
        
           '--------flush into Temp_Tabel-------------------------------
            con.connectionstring = strcn.Connection
            con.Open
            
            Temp_Table_Helper.Open "select m_code,s_code,(select s_name=isnull(s_name,'') from test_info_sub Where test_info_sub.s_code = pat_info_sub1.s_code and test_info_sub.m_code=pat_info_sub1.m_code and pat_id='" + txtPat_ID + "') as s_name,test_rate,delv_dt,type,unique_id from pat_info_sub1 where pat_id='" + txtDummy_Pat_ID.text + "'", con
            
            'MsgBox Temp_Table_Helper.RecordCount
              While Temp_Table_Helper.EOF = False
                    Temp_Table.AddNew
                                                            
                    Temp_Table!m_code = Temp_Table_Helper!m_code
                    Temp_Table!s_code = Temp_Table_Helper!s_code
                    Temp_Table!s_name = Temp_Table_Helper!s_name
                    Temp_Table!test_rate = Temp_Table_Helper!test_rate
                    Temp_Table!Delv_DTM = Temp_Table_Helper!Delv_Dt
                    Temp_Table!Type = Temp_Table_Helper!Type
                    Temp_Table_Helper.MoveNext
              Wend
                
            DataGrid1.Refresh
            Temp_Table_Helper.Close
            con.Close
           
           
           '---------------------------------------------------------
                     
               '`````````````to show DISCOUNT from pat_info_sub3``````
               Adodc6.connectionstring = strcn.Connection
               Adodc6.RecordSource = "exec Pat_Info_SELECT 11,'" & txtDummy_Pat_ID.text & "'"
               Adodc6.Refresh
    
               If Adodc6.Recordset.RecordCount > 0 Then
               Dim strchkpaid As String
                nbrDisc.text = "0"
                
                nbrTot_Disc = Adodc6.Recordset!disc
                strchkpaid = Adodc6.Recordset!paid
                'MsgBox strchkpaid
                    If Trim(strchkpaid) = "True" Then
                    ChkPaid.value = 1
                    ChkPaidVal = "1"
                    Else
                    ChkPaid.value = 0
                    ChkPaidVal = "0"
                    End If
               End If
               '```````````````````````````````````````````````````````
               
               '`````````````to show REFERENCE_TYPE from pat_info_MAIN``````
               Adodc6.connectionstring = strcn.Connection
               Adodc6.RecordSource = "exec Pat_Info_SELECT 1,'" + txtDummy_Pat_ID.text + "'"
               Adodc6.Refresh
    
               If Adodc6.Recordset.RecordCount > 0 Then
               Dim strRefer_Type1 As String
               
                strRefer_Type1 = Adodc6.Recordset!refer_type
                    If strRefer_Type1 = 1 Then
                    Chkrefer_type.value = 1
                    strRefer_Type1 = "1"
                    Else
                    Chkrefer_type.value = 0
                    strRefer_Type1 = "0"
                    End If
               End If
               '``````````````````````````````````````````````````````
               
               '*************for flush doctor ID and name ****************
               Adodc12.connectionstring = strcn.Connection
               Adodc12.RecordSource = "exec Pat_Info_SELECT 7,'" + txtDummy_Pat_ID.text + "'"
               
               Adodc12.Refresh
               If Adodc12.Recordset.RecordCount > 0 Then
               
                   txtRefer_Code = Adodc12.Recordset!Refer_code
                'MsgBox txtRefer_Code
               
               End If
               
               '*************for flush Consultant ID and name ****************
               Adodc12.connectionstring = strcn.Connection
               Adodc12.RecordSource = "exec Pat_Info_SELECT 15,'" + txtDummy_Pat_ID.text + "'"
               
               Adodc12.Refresh
               If Adodc12.Recordset.RecordCount > 0 Then
               
                   txtCons_Code = Adodc12.Recordset!Cons_Code
                'MsgBox txtRefer_Code
               
               End If
               
'              '======DOCTOR NAME FROM DOCTOR_INFO_NEW=============
               Adodc13.connectionstring = strcn.Connection
               Adodc13.RecordSource = "exec Pat_Info_SELECT 6,'" + txtDummy_Pat_ID.text + "'"
               
               Adodc13.Refresh
               If Adodc13.Recordset.RecordCount > 0 Then
               
                  txtDoc_Name = Adodc13.Recordset!doc_name
                  txtDoc_Addr = Adodc13.Recordset!addr
               End If
               '=====================END===========================
               ',,,,,,,,,,,,,,for get registered doctor,,,,,,,,,,,
               Dim My_Rst As New ADODB.Recordset
               con.connectionstring = strcn.Connection
               con.Open
               Set My_Rst.ActiveConnection = con
               My_Rst.Open "exec Pro_FLUSH1 1,'" & Trim(txtRefer_Code.text) & "'", con
               If My_Rst.EOF = False Then
                  
                    txtDoc_Name.text = My_Rst!doc_name
                    txtDoc_Addr.text = My_Rst!addr
               Else
                    txtDoc_Name.ForeColor = vbBlack
                    txtDoc_Addr.ForeColor = vbBlack
               End If
               My_Rst.Close
               con.Close
               
               ',,,,,,,,,,,,,,for get MExecutive,,,,,,,,,,,
               Dim ME_Rst As New ADODB.Recordset
               con.connectionstring = strcn.Connection
               con.Open
               Set ME_Rst.ActiveConnection = con
               ME_Rst.Open "exec Pro_FLUSH3 1,'" & Trim(txtEmail.text) & "'", con
               If ME_Rst.EOF = False Then
                  
                    txtMEName.text = ME_Rst!Emp_Name
'                    txtDoc_Addr.text = My_Rst!addr
               Else
                    txtMEName.ForeColor = vbBlack
'                    txtDoc_Addr.ForeColor = vbBlack
               End If
               ME_Rst.Close
               con.Close
                
               ',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
              
               '***************end****************************************
                
         Else
           txtPat_Name = ""
           ComSex = "Male"
           txtAge = ""
           txtRefer_Code = ""
           txtDegree = ""
           txtAddr = ""
           txtPhone = ""
           txtFax = ""
           txtEmail = ""
           Dt.value = Now
           Delv_Dt.value = Now
           nbrVAT_Amt = 0
           nbrAdv = 0
           nbrDisc = 0
           nbrTot_Disc = 0
           nbrDisc_Per = 0
           nbrDue = ""
           nbrNet_Amount = ""
           
           nbrTest_Rate = ""
           nbrTotal = ""
           ChkPaid.value = 0
           Delv_TM.value = Now
           Chkrefer_type.value = 0
        End If
        
'++++++++++for count TOTAL_RATE from Temp_Table+++++++++
        If Temp_Table.RecordCount > 0 Then
        Total_Rate = 0
        Temp_Table.MoveFirst
        While Temp_Table.EOF = False
                Total_Rate = Total_Rate + Temp_Table!test_rate

        Temp_Table.MoveNext
        Wend
        nbrTotal = Total_Rate
        End If
'++++++++++End count TOTAL_RATE from Temp_Table+++++++++
        
'=========count total test=============================
        Total_Test = 0
        Total_Test = Temp_Table.RecordCount
        Me.nbrTot_Test = Total_Test
'=========End count total test========================

'>>>>>>>>>>>>>>>>to show total advance>>>>>>>>>>>>>>>>>>>>>>
    Adodc7.connectionstring = strcn.Connection
    Adodc7.RecordSource = "exec Pro_FLUSH 3,'" & txtDummy_Pat_ID.text & "'"
    Adodc7.Refresh
    If Adodc7.Recordset.RecordCount > 0 Then
        nbrAdv.text = Adodc7.Recordset!adv_sum
        nbrTotCollect_Fee.text = Adodc7.Recordset!Coll_sum
    End If
'<<<<<<<<<<<<End show total advance<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    DataGrid1.Columns(0).Width = 450.1418
    DataGrid1.Columns(1).Width = 810.1418
    DataGrid1.Columns(2).Width = 3825.071
    DataGrid1.Columns(3).Width = 1110.047
    DataGrid1.Columns(4).Width = 2100
    DataGrid1.Columns(5).Width = 800

nbrAdv_Pay.SetFocus

    
End Sub

Private Sub txtPat_Name_GotFocus()
txtPat_Name.BackColor = &HFFFFC0
End Sub

Private Sub txtPat_Name_LostFocus()
    txtPat_Name.BackColor = vbWhite
    txtPat_Name.text = StrConv(txtPat_Name.text, vbProperCase)
End Sub

'Private Sub txtPat_Name_Change()
'
'    On Error Resume Next
'
'    txtPat_Name = StrConv(txtPat_Name, vbProperCase)
'
'    On Error GoTo 0
'
'End Sub

Private Sub txtPat_Name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       SendKeys Chr(9)
    txtPat_Name.text = StrConv(txtPat_Name.text, vbProperCase)
End If
End Sub

Private Sub txtPhone_GotFocus()
txtPhone.BackColor = &HFFFFC0
End Sub

Private Sub txtPhone_LostFocus()
txtPhone.BackColor = vbWhite
End Sub

Private Sub txtRefer_Code_GotFocus()
txtRefer_Code.BackColor = &HFFFFC0
End Sub

Private Sub txtRefer_Code_LostFocus()
On Error GoTo err_sub
    txtRefer_Code.BackColor = vbWhite
    
    txtM_Code.TabStop = True
    If Trim(txtRefer_Code) = "" Then Exit Sub
    'MsgBox "Patient1"
    Doc_List_MODE = "frmPatient_Info"

       If Trim(txtRefer_Code.text) = "0" Then
       
            If Trim(txtPat_ID.text) <> "" Then
                
                If u_id <> "md" Then
                MsgBox "If you want to any change you should contact to Managing Director.., Your ID saved..", vbCritical
                txtRefer_Code = ""
                Exit Sub
                End If
                NdocMode = "0"
                frmDoctor_Info_New.txtPat_ID = txtPat_ID
            End If
            
            If Trim(txtPat_ID.text) = "" Then
                NdocMode = "1"
                frmDoctor_Info_New.txtPat_ID = "0"
            End If
       
       frmDoctor_Info_New.Show vbModal 'for new unknown doctor
       
       Else
               Adodc2.connectionstring = strcn.Connection
               
               Adodc2.RecordSource = "exec Pro_FLUSH1 1,'" & Trim(txtRefer_Code.text) & "'"
               Adodc2.Refresh
               
               
                'MsgBox "Patient2"
               If Adodc2.Recordset.RecordCount > 0 Then
                   txtDoc_Name.text = Adodc2.Recordset!doc_name
                   txtDoc_Addr.text = Adodc2.Recordset!addr
                   txtM_Code.TabStop = True
               Else
               
                   'MsgBox "Patient3"
                   txtM_Code.TabStop = False
                   frmDoc_List.Show vbModal
                   Exit Sub
               End If
       End If
    Exit Sub
    
err_sub:
    MsgBox Err.Description
    
End Sub

Private Sub txtS_Code_GotFocus()
txtS_Code.BackColor = &HFFFFC0
End Sub

Private Sub txtS_Code_LostFocus()

On Error Resume Next
    
    txtS_Code.BackColor = vbWhite
    
    If Trim(txtS_Code) = "" Then Exit Sub

    Adodc4.connectionstring = strcn.Connection
    Adodc4.RecordSource = "exec  sp_found '" + txtM_Code + "','" + txtS_Code + "'"
    Adodc4.Refresh

    If Adodc4.Recordset.Fields(0) = "N" Then
        Test_List_Mode = "frmPatient_Info_S"
        txtS_Name = ""
        nbrTest_Rate = 0
        txtType.text = ""
        txtS_Code = ""
        frmTest_List.Show vbModal
    Else
        If Len(Trim(txtM_Code.text)) = 0 Then
            MsgBox "Group Code mandatory"
            txtM_Code.SetFocus
            Exit Sub
        End If
        txtS_Name = Adodc4.Recordset.Fields(0)
        nbrTest_Rate = Adodc4.Recordset.Fields(1)
        txtType.text = Adodc4.Recordset.Fields(2)
    End If
        
End Sub

Public Sub Temp_rst()
    '--------------------------------------------
    Set Temp_Table = New ADODB.Recordset
    With Temp_Table
        .Fields.Append "m_code", adVarChar, 2
        .Fields.Append "s_code", adVarChar, 3
        .Fields.Append "s_name", adVarChar, 60
        .Fields.Append "test_rate", adDouble
        .Fields.Append "Delv_DTM", adVarChar, 26
        .Fields.Append "type", adVarChar, 10
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set DataGrid1.DataSource = Temp_Table
    
    DataGrid1.ReBind
    DataGrid1.Refresh
    
    DataGrid1.Columns(0).Width = 450.1418
    DataGrid1.Columns(1).Width = 810.1418
    DataGrid1.Columns(2).Width = 3825.071
    DataGrid1.Columns(3).Width = 1110.047
    DataGrid1.Columns(4).Width = 2100
    DataGrid1.Columns(5).Width = 800
    
    
End Sub

Private Sub Select_Unique_ID()
    If Len(Trim(txtPat_ID.text)) = 0 Then Exit Sub
    If Len(Trim(txtM_Code.text)) = 0 Then Exit Sub
    If Len(Trim(txtS_Code.text)) = 0 Then Exit Sub
    
    Adodc8.connectionstring = strcn.Connection
    Adodc8.RecordSource = "exec pro_flush_unique_id 1,'" + txtPat_ID + "','" + txtM_Code + "','" + txtS_Code + "'"
    
    Adodc8.Refresh
    If Adodc8.Recordset.RecordCount > 0 Then
    nbrUnique_id = Adodc8.Recordset!unique_id
    Else
    nbrUnique_id = ""
    End If
End Sub

Private Sub Auto_No()

    
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    My_Rst.Open "select count(pat_id)+1 from pat_info_main", con
    If IsNull(My_Rst.Fields(0)) = False Then
       txtPat_ID = BoothN + pad("l", 9, My_Rst.Fields(0), "0")
    End If
    My_Rst.Close
    con.Close
       

End Sub

Private Sub nbrVAT_Per_Change()
    nbrVAT_Amt = Val(nbrTotal) * Val(nbrVAT_Per) / 100
    nbrTotal_Amt = Val(nbrTotal) + Val(nbrVAT_Amt)
End Sub

Private Sub InsPat_Info_Sub2()
    If Trim(brAdv_Pay) = 0 Or Trim(nbrAdv_Pay) = "" Then Exit Sub
    
    
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_PAT_INFO_SUB2 'I'," + "0" + _
    "," + nbrAdv_Pay + ",'" + u_id + "','" + CDate_TM + "','" + "" + "'"
    cmd.Execute
    con.Close
End Sub

Private Sub InsPat_Info_Sub2_T()
    If Trim(nbrAdv_Pay) = "" Then
        nbrAdv_Pay = "0"
    
    
    End If
   
    
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_PAT_INFO_SUB2 'I'," + StPat_ID + _
    "," + nbrAdv_Pay + ",'" + u_id + _
    "','" + CDate_TM + _
    "'," + nbrCollect_Fee.text + _
    "," + "ADV" + _
    ",'" + Format(Dt, "yyyy-mm-dd") + _
    "','" + CDate_TM + _
    "','" + Format(CDate_TM, "yyyy-mm-dd") + _
    "','" + "" + "'"
'    MsgBox cmd.CommandText
    cmd.Execute
    con.Close
    
End Sub

Private Sub InsPat_Info_Sub2_T1()

    If Trim(brAdv_Pay) = 0 Or Trim(nbrAdv_Pay) = "" Then Exit Sub
           
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_PAT_INFO_SUB2 'I'," + txtPat_ID + _
    "," + nbrAdv_Pay + _
    ",'" + u_id + _
    "','" + CDate_TM + _
    "'," + Trim(nbrCollect_Fee.text) + _
    "," + "DUE" + _
    ",'" + Format(CDate_TM5, "yyyy-mm-dd") + _
    "','" + Format(CDate_TM8, "yyyy-mm-dd hh:mm") + _
    "','" + Format(CDate_TM10, "yyyy-mm-dd") + _
    "','" + "" + "'"
    cmd.Execute
    con.Close
    
End Sub

Private Sub InsDoc_info_new()
    Dim strRefer_Code As String
    Dim StrDoc_Name As String
    Dim strAddr As String
    Dim strPhone As String
    Dim strFax As String
    Dim strEmail As String
    Dim StrUid As String
    Dim strDoc_Date As String
    
    Adodc15.connectionstring = strcn.Connection
    
    Adodc15.RecordSource = "exec New_Doc_Select 2,'','" & u_id & "'"
    Adodc15.Refresh
    If Adodc15.Recordset.RecordCount > 0 Then
        strRefer_Code = Adodc15.Recordset!pat_id
        
    
    '-------UPDATE DOCTOR ID into doctor_info_new------------------------
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec pro_DOCTOR_INFO_NEW2 'U','" & StPat_ID & "','" & u_id & "'"

    cmd.Execute
    con.Close
    '-----------------------------------------------------------
    '>>>>>>>>>>>>>>>>>>
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    cmd.CommandText = "exec PAT_INFO_MAIN_U 'U'," & StPat_ID & ""
    
    cmd.Execute
    con.Close
    '>>>>>>>>>>>>>>>>>>>
    
    End If
End Sub

Private Sub InsD_TM()


    '++++++for insert Current Date and Time++++++++++++++
    Dim StrCdt As String
    Dim StrCtime As String
     
    StrCdt = Trim(Format(Dt, "yyyy-mm-dd"))
    StrCtime = Trim(Format(DT_TM, "hh:mm"))
    CDate_TM = StrCdt + Space(1) + StrCtime
    CDate_TM10 = StrCdt
   '++++++++++end+++++++++++++++++++++++++++++++++++++++
End Sub

Private Sub Sel_Refer_Type()
    
    If Chkrefer_type.value = 1 Then
        StrRefer_Type = "1"
    End If
    
    If Chkrefer_type.value = 0 Then
        StrRefer_Type = "0"
    End If
End Sub

Private Sub Search_Refer_Code() 'search again refer_code for update refer_code/delete from doctor_info_new
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Doc_SELECT 4,'" + txtPat_ID.text + "'", con
    If My_Rst.EOF = False Then
        Del_Doc = My_Rst!Refer_code
        
    End If
    con.Close
End Sub

Private Sub Del_New_Doc()

    If Del_Doc <> "" Then ''''delete from doctor_info_new
       'MsgBox "del"
        con.connectionstring = strcn.Connection
        con.Open
        Set cmd.ActiveConnection = con
        cmd.CommandText = "exec delete_all 1," + txtPat_ID + ""
        cmd.Execute
        con.Close
        
       End If
End Sub

Private Sub Flush_VAT_Per()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec pro_name_SELECT '19',''", con
    If My_Rst.EOF = False Then
        nbrVAT_Per.text = My_Rst!vat_per
    End If
    
    con.Close
End Sub

Private Sub Make_Pat_ID1()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Make_Pat_ID1 '" & Chkrefer_type.value & "'", con
    If My_Rst.EOF = False Then
        Strpat_id1 = My_Rst!pat_id1
        Strpat_MY = My_Rst!pat_my
'        MsgBox Strpat_id1
    End If
    
    con.Close
    
End Sub

Private Sub Make_Pat_ID1_U()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Make_Pat_ID_U '" & Chkrefer_type.value & "'", con
    If My_Rst.EOF = False Then
        Strpat_id1 = My_Rst!pat_id1
        Strpat_MY = My_Rst!pat_my
'        MsgBox Strpat_id1
    End If
    
    con.Close
    
End Sub

Private Sub Search_Patient_Type()

    StrRow_Count = "1"
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Search_Pat_Type 1,'" & txtPat_ID1.text & "'", con
    If My_Rst.EOF = False Then
    
        StrRow_Count = My_Rst!Row_Count
        'MsgBox StrRow_Count
    End If
    
    con.Close
    
End Sub

Private Sub Srch_Pat_ID()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Search_Pat_ID 1,'" & txtPat_ID1.text & "','" & StrPat_Type & "'", con
    If My_Rst.EOF = False Then
        IntPat_ID = My_Rst!pat_id2
  '      MsgBox IntPat_ID
    End If
    con.Close
End Sub

Private Sub Srch_Pat_ID1()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Search_Pat_ID1 1,'" & txtPat_ID1.text & "'", con
    If My_Rst.EOF = False Then
        IntPat_ID = My_Rst!pat_id2
 '       MsgBox IntPat_ID
    End If
    con.Close
    
End Sub

Private Sub Flush_Pat_ID()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Pat_Info_SELECT1 1,'" & txtPat_ID1.text & "'", con
    If My_Rst.EOF = False Then
        IntPat_ID = My_Rst!pat_id
'        MsgBox IntPat_ID
    End If
    con.Close
    
End Sub

Private Sub GATE_DT()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec CR_Date", con
    If My_Rst.EOF = False Then
        Dt.value = My_Rst!crDate
        DT_TM.value = My_Rst!crDate
'        MsgBox IntPat_ID
    End If
    con.Close
    
End Sub

Private Sub Cal_Dis()

    
    DblDisc = Val(nbrTotal_Amt) * Val(nbrDisc_Per) / 100

End Sub
Private Sub Del_False_New_Doc()

    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
    My_Rst.Open "exec Del_Doc_New 1,'" & "0" & "','" & u_id & "'", con

    con.Close
    
End Sub

Private Sub txtS_Name_GotFocus()
txtS_Name.BackColor = &HFFFFC0
End Sub

Private Sub txtS_Name_LostFocus()

    txtS_Name.BackColor = vbWhite
    
End Sub


Public Sub PrintReport()

On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    Dim StrPat_ID As String
   
    Set rscashmaster = New ADODB.Recordset
    
    
    If frmPatient_Info.txtPat_ID = "" Then
            StrPat_ID = StPat_ID
          Else
            StrPat_ID = frmPatient_Info.txtPat_ID
    End If

If rscashmaster.State <> 0 Then rscashmaster.Close
        rscashmaster.Open "exec Rpt_pat_info '" & StrPat_ID & "'", strcn.Connection
    
    
    
    If rscashmaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\CashMemo1.RPT"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
        



             
                   
'        Set objReportFF = objReportFormulaFieldDefinations.Item(1)
'            objReportFF.Text = "'" + parseQuotes(txtWords.Text) + " '"
'
'            Set objReportFF = objReportFormulaFieldDefinations.Item(2)
'
'            objReportFF.Text = "'" + parseQuotes(txtUserName.Text) + " '"
'
''            -------------------Add Discunt------------------
'           If Val(txtTotalDiscount.Text) <> 0 And Val(txtDiscount.Text) <> 0 Then
'           Set objReportFF = objReportFormulaFieldDefinations.Item(3)
'            objReportFF.Text = "'" + parseQuotes(txtDiscount.Text) + " '"
'
'            Set objReportFF = objReportFormulaFieldDefinations.Item(4)
'
'            objReportFF.Text = "'" + parseQuotes(txtTotalDiscount.Text) + " '"
'
'
'            Set objReportFF = objReportFormulaFieldDefinations.Item(5)
'
'            objReportFF.Text = "'" + "Special Discount" + " '"
'
'            Set objReportFF = objReportFormulaFieldDefinations.Item(6)
'
'            objReportFF.Text = "'" + "%" + " '"

'End If
'-------------End Add Discunt-------------------
        objReportDatabaseTable.SetPrivateData 3, rscashmaster
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
        If Tracer = 0 Then
        objReport.Preview "Payment Report", , , , , 16777216 Or 524288 Or 65536
        Else
        objReport.PrintOut
        End If
        
        
        Set objReport = Nothing
        Set objReportDatabase = Nothing
        Set objReportDatabaseTables = Nothing
        Set objReportDatabaseTable = Nothing
    Exit Sub

ErrH:

    Select Case Err.Number
        Case 20545
            MsgBox "Request cancelled by the user", vbInformation, "Printing Cancel Information"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Printing Cancel Information"
    End Select
    
    
   
End Sub

Public Sub printReport1()

On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    Dim StrPat_ID As String
   
    Set rscashmaster = New ADODB.Recordset
    
    
    If frmPatient_Info.txtPat_ID = "" Then
            StrPat_ID = StPat_ID
          Else
            StrPat_ID = frmPatient_Info.txtPat_ID
    End If

'            Report14.DiscardSavedData
'            RS.Open "exec Rpt_pat_info '" & StrPat_ID & "'", strcn.Connection
'            Report14.Database.SetDataSource RS
'            Report14.Text29.SetText frmPatient_Info.txtRefer_Code.Text
'            CRViewer1.ReportSource = Report14
           

    

If rscashmaster.State <> 0 Then rscashmaster.Close
        rscashmaster.Open "exec Rpt_pat_info '" & StrPat_ID & "'", strcn.Connection
    
    
    
    If rscashmaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\CashMemo1.RPT"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
        
'-------------End Add Discunt-------------------
        objReportDatabaseTable.SetPrivateData 3, rscashmaster
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
'        If Tracer = 0 Then
'        objReport.Preview "Cash Memo Report", , , , , 16777216 Or 524288 Or 65536
'        Else
        objReport.PrintOut (False)
'         objReport.ProgressDialogEnabled
'        End If
        
        
        Set objReport = Nothing
        Set objReportDatabase = Nothing
        Set objReportDatabaseTables = Nothing
        Set objReportDatabaseTable = Nothing
        
        Call cmdNew_Click
    Exit Sub

ErrH:

    Select Case Err.Number
        Case 20545
            MsgBox "Request cancelled by the user", vbInformation, "Printing Cancel Information"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Printing Cancel Information"
    End Select
    
   
End Sub

Public Function LwrToUpprCase(ByVal edit$) As String
    Dim A$, b$, c$, i%
    On Error Resume Next

    If Len(Trim(edit$)) = 0 Then
        LwrToUpprCase = edit$
        Exit Function
    End If

    edit$ = edit$ + Space$(3)       'end of word match ???
    Mid$(edit$, 1, 1) = UCase$(Mid$(edit$, 1, 1)) 'cap first char
    If Mid$(edit$, 1, 3) = "Po " Then
        Mid$(edit$, 1, 3) = "PO " 'PO Box cap
    End If
    b$ = "NwNeSwSeIi"       'NW NE SW SE II
    If Mid$(edit$, 1, 2) = "Mc" Then
        Mid$(edit$, 3, 1) = UCase$(Mid$(edit$, 3, 1))
    End If
    If Mid$(edit$, 1, 1) > Chr$(47) And Mid$(edit$, 1, 1) < Chr$(58) Then
        Mid$(edit$, 2, 1) = UCase$(Mid$(edit$, 2, 1)) 'nos 0 - 9
    End If
    c$ = " >-./#&,(!@$%^*+=\|" & Chr(34)
    If InStr(c$, Mid$(edit$, 1, 1)) > 0 Then
        Mid$(edit$, 2, 1) = UCase$(Mid$(edit$, 2, 1))
    End If
    For i% = 2 To Len(edit$)
        If InStr(b$, Mid$(edit$, i%, 2)) > 0 Then
            If Mid$(edit$, i% - 1, 1) = " " And Mid$(edit$, i% + 2, 1) = " " Then
                Mid$(edit$, i%, 2) = UCase$(Mid$(edit$, i%, 2))
            End If
            i% = i% + 2
        End If
        If Mid$(edit$, i%, 4) = " of " Then 'dont cap word "of"
            i% = i% + 2
            GoTo skipcapchar
        ElseIf Mid$(edit$, i%, 5) = " and " Then 'dont cap word "and"
            i% = i% + 3
            GoTo skipcapchar
        
        ElseIf Mid$(edit$, i%, 5) = " c/o " Then 'dont cap abbrev c/o
            i% = i% + 3
            GoTo skipcapchar
        End If
        If InStr(c$, Mid$(edit$, i%, 1)) > 0 And i% < Len(edit$) Then
            Mid$(edit$, i% + 1, 1) = UCase(Mid$(edit$, i% + 1, 1))
        End If
        If Mid$(edit$, i%, 1) = Chr$(39) And Mid$(edit$, i% + 1, 1) <> "s" Then
            Mid$(edit$, i% + 1, 1) = UCase$(Mid$(edit$, i% + 1, 1)) 'O'Malley's
        End If
        If Mid$(edit$, i%, 2) = "Mc" Then
            Mid$(edit$, i% + 2, 1) = UCase$(Mid$(edit$, i% + 2, 1))
        End If
        If InStr("0123456789", Mid$(edit$, i%, 1)) > 0 Then 'nos 0-9
            A$ = Mid$(edit$, i% + 1, 2)
            If A$ <> "st" And A$ <> "th" And A$ <> "nd" And A$ <> "rd" Then
                Mid$(edit$, i% + 1, 1) = UCase$(Mid$(edit$, i% + 1, 1)) 'no's 0-9
            End If
        End If
skipcapchar:
    Next i%
    edit$ = RTrim$(edit$) 'trim space added earlier

    LwrToUpprCase = edit$
End Function


Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If KeyAscii = 13 And txtSearch = "" Then
'SendKeys Chr(9)
nbrDisc.SetFocus
End If
If txtSearch <> "" Then
lvItemSearch.Show vbModal
txtSearch.text = ""
End If
End If
End Sub

