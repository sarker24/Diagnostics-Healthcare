VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form rMicrobiology1 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Lab Report Format [MICROBIOLOGY SPECIAL]"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   DrawWidth       =   2
   Icon            =   "rMicrobiology1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDel_Temp 
      Caption         =   "D e l e t e"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   10260
      TabIndex        =   36
      Top             =   3720
      Width           =   360
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
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
      Left            =   8070
      TabIndex        =   12
      Top             =   3360
      Width           =   1050
   End
   Begin VB.TextBox txtOrg_Iso1 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2250
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2400
      Width           =   6810
   End
   Begin VB.TextBox txtTest_Name 
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   2220
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   3720
      Width           =   7950
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   5160
      Top             =   300
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "Adodc7"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "rMicrobiology1.frx":000C
      Height          =   990
      Left            =   3720
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   1746
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
   Begin VB.TextBox txtSpecimen 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2250
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1830
      Width           =   3390
   End
   Begin VB.TextBox txtM_Code 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3735
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   345
   End
   Begin VB.TextBox txtS_Code 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2250
      TabIndex        =   3
      Top             =   1455
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox txtS_Name 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3750
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1455
      Visible         =   0   'False
      Width           =   5325
   End
   Begin VB.TextBox txtUsed_tech 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2250
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2970
      Width           =   6795
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
      Left            =   9135
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1050
   End
   Begin VB.ComboBox ComTest_Title 
      Height          =   315
      Left            =   2220
      TabIndex        =   10
      Top             =   3315
      Width           =   3225
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
      Left            =   6435
      TabIndex        =   17
      Top             =   10305
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
      TabIndex        =   19
      Top             =   10305
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
      TabIndex        =   20
      Top             =   10305
      Width           =   1050
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   7485
      TabIndex        =   18
      Top             =   10305
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
      Left            =   4335
      TabIndex        =   15
      Top             =   10305
      Width           =   1050
   End
   Begin VB.TextBox txtNote 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   9210
      Width           =   10110
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
      Left            =   5385
      TabIndex        =   16
      Top             =   10305
      Width           =   1050
   End
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   9270
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
      Caption         =   "9-S_NAME"
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
      Left            =   9270
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
      Caption         =   "Adodc8"
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
      Left            =   9270
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
      Caption         =   "3-pat_ID"
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
      Left            =   9270
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   9270
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
      Caption         =   "2-Show Advance"
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
      Left            =   9270
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
      Caption         =   "6-S_code lost_focus"
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
      Left            =   9270
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
      Caption         =   "Adodc5"
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
      Left            =   9270
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
      Caption         =   "Adodc4"
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
   Begin MSComCtl2.DTPicker Dt 
      Height          =   285
      Left            =   7875
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      _Version        =   393216
      Format          =   66781185
      CurrentDate     =   37114
   End
   Begin VB.TextBox txtSN 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   450
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   10035
      Visible         =   0   'False
      Width           =   10140
   End
   Begin RichTextLib.RichTextBox txtOrg_Iso 
      Height          =   255
      Left            =   2220
      TabIndex        =   33
      Top             =   2640
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"rMicrobiology1.frx":0021
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox txtN_Exam 
      Height          =   315
      ItemData        =   "rMicrobiology1.frx":009A
      Left            =   2220
      List            =   "rMicrobiology1.frx":00AD
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2100
      Width           =   3405
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1995
      Left            =   2190
      TabIndex        =   35
      Top             =   6960
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3519
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   27
      FormatLocked    =   -1  'True
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995.024
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPat_ID1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2250
      MaxLength       =   10
      TabIndex        =   0
      ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
      Top             =   1065
      Width           =   1290
   End
   Begin VB.TextBox txtPat_ID 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2250
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1065
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special Note"
      Height          =   195
      Left            =   450
      TabIndex        =   32
      Top             =   9750
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Organism isolated"
      Height          =   195
      Left            =   495
      TabIndex        =   31
      Top             =   2610
      Width           =   1245
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature of Examination"
      Height          =   195
      Left            =   495
      TabIndex        =   30
      Top             =   2160
      Width           =   1560
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Specimen"
      Height          =   195
      Left            =   495
      TabIndex        =   29
      Top             =   1845
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Impression"
      Height          =   195
      Left            =   480
      TabIndex        =   28
      Top             =   8940
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Used Technology"
      Height          =   195
      Left            =   495
      TabIndex        =   27
      Top             =   2910
      Width           =   1260
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   7320
      TabIndex        =   26
      Top             =   1050
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID"
      Height          =   195
      Left            =   495
      TabIndex        =   25
      Top             =   1050
      Width           =   705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Code"
      Height          =   195
      Left            =   495
      TabIndex        =   24
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   3240
      TabIndex        =   23
      Top             =   1440
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MICROBIOLOGY SPECIAL"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   510
      TabIndex        =   22
      Top             =   180
      Width           =   2505
   End
End
Attribute VB_Name = "rMicrobiology1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp_Table1 As New ADODB.Recordset
Dim Temp_Table_Helper1 As New ADODB.Recordset

Dim Strpat_id1 As String
Dim StrRow_Count As String
Dim StrPat_Type As String
Dim IntPat_ID As Double

Private Sub cmdClear_Click()
    Temp_rst1
    
    ComTest_Title.text = ""
    txtTest_Name.text = ""
'    txtTest_Result.Text = ""
    txtNote.text = ""
    txtSN.text = ""
    txtPat_ID1 = ""
    txtPat_ID = ""
    txtOrg_Iso.text = ""
    txtSpecimen = ""
    txtUsed_tech = ""
    ComTest_Title.Clear
    If DataGrid1.Visible = True Then
        DataGrid1.Visible = False
    End If
    txtPat_ID1.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDel_Temp_Click()

        Dim Strmsg As String
        Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
        If Strmsg = vbYes Then
        DelReport_All_TempRst1
        Temp_Table1.Delete
        ComTest_Title = ""
        txtTest_Name = ""
        End If


End Sub

Private Sub cmdDelete_Click()
    Temp_rst1
    If cmdSave.Enabled = False Then Exit Sub
    Dim Strmsg As String
    Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
    If Strmsg = vbYes Then
    'Del_All_Report_All_TempRst1
    Del_Report
    Clearscreen
    txtPat_ID1.SetFocus
    End If
End Sub
Private Sub cmdDelete_TempTable1_Click()

    If ComTest_Name = "" Then Exit Sub
    If cmdSave.Enabled = False Then Exit Sub
  '  If Temp_Table1.RecordCount <= 0 Then Exit Sub
    
    If Trim(ComTest_Name.text) = "" Then
        MsgBox "You didn't select the the Test Name"
        DataGrid2.SetFocus
        Exit Sub
    Else
        Dim Strmsg As String
        Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
        If Strmsg = vbYes Then
'        DelReport_All_TempRst1
   '     Temp_Table1.Delete
        ComTest_Title.text = ""
        txtTest_Name.text = ""
        txtTest_Result = ""
'        txtUnit = ""
'        txtRef_Range = ""
        End If
        
    End If
End Sub

Private Sub CmdInsert_Click()

'If Trim(ComTest_Name) = "" Then Exit Sub
'----------------check--------
Dim Check As Integer
Check = 0
If Temp_Table1.RecordCount > 0 Then
    Temp_Table1.MoveFirst
    
        While Temp_Table1.EOF = False
                
            If Temp_Table1!Test_Name = ComTest_Title Then
                Check = 1
            End If
    Temp_Table1.MoveNext
        Wend
    If Check = 1 Then
        MsgBox "This Test Name already exists"
        Check = 0
        ComTest_Title.SetFocus
        Exit Sub
    End If
'    Temp_Table.MoveFirst
End If

'--------------end check-----

'+++to insert into TEMPORARY RECORDSET "Temp_rst1"++++
        Temp_Table1.AddNew
        Temp_Table1!Test_Name = ComTest_Title
        Temp_Table1!Test_result = txtTest_Name
        'Temp_Table1!unit = txtUnit
        DataGrid2.Refresh
'+++++++++++++++++++++++++++++++++++++++

    DataGrid2.Columns(0).Width = 1000
    
    ComTest_Title.SetFocus
    
    ComTest_Title = ""
    txtTest_Name = ""


DataGrid2.Columns(0).Width = 2970.142
DataGrid2.Columns(1).Width = 5000

End Sub

Private Sub CmdPreview_Click()
    CRViewer1_MODE = 46
    Viewer1.Show vbModal
End Sub

Private Sub cmdPrint_Click()
    '==========direct print==========================
            
            Dim Report46 As New Microbiology1
            Dim StrPat_ID As String
        
            
            Dim strM_Code As String
            Dim strS_Code As String
            
            StrPat_ID = rMicrobiology1.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rMicrobiology1.txtM_Code
            strS_Code = rMicrobiology1.txtS_Code
            
                       '--------------------------------------------------------------------
            Report46.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report46.FormulaFields.Item(2).text = Chr(34) & "Patient ID" & Chr(34)
            Report46.FormulaFields.Item(3).text = Chr(34) & "Reeived Date" & Chr(34)
            Report46.FormulaFields.Item(4).text = Chr(34) & "Delivered Date" & Chr(34)
            Report46.FormulaFields.Item(5).text = Chr(34) & "Patient Name" & Chr(34)
            Report46.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report46.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report46.FormulaFields.Item(8).text = Chr(34) & "Refd. by" & Chr(34)
            '--------------------------------------------------------------------
            Report46.FormulaFields.Item(9).text = Chr(34) & "Specimen" & Chr(34)
            Report46.FormulaFields.Item(10).text = Chr(34) & "Nature of Exam" & Chr(34)
            Report46.FormulaFields.Item(11).text = Chr(34) & "1. Organism Isolated" & Chr(34)
            Report46.FormulaFields.Item(12).text = Chr(34) & "2. Sensitivity Test" & Chr(34)
            Report46.FormulaFields.Item(13).text = Chr(34) & "Checked By" & Chr(34)
 
'            Report46.Text2.SetText Trim(rMicrobiology1.txtTest_Name.Text)
'            Report46.Text3.SetText Trim(rMicrobiology1.txtTest_Result.Text)

            Call Flush_Doc_Name
            Report46.Text2.SetText StDoc_Name

            
            Report46.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report46.Database.SetDataSource rs
            
            Report46.PrintOut (False)
            rs.Close
            
            Call cmdClear_Click
            txtPat_ID1.SetFocus
            
    '====================================
End Sub

Private Sub cmdSave_Click()
'-----validation check---------------------
    If txtPat_ID1 = "" Then
        MsgBox "Patient ID mandatory"
        txtPat_ID1.SetFocus
        Exit Sub
    End If
    
'-----end validation check--------------------------------------------------

''\\\\\\\\\\INSERT and UPDATE from Temp_rst2\\\\\\\\\\\\\
    Adodc2.connectionstring = strcn.Connection
    Adodc2.RecordSource = "Report_All_SELECT3 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then
        Del_All_Report_All_TempRst1
        InsReport_All_TempRst1
        MsgBox "Updated1"
    Else
        InsReport_All_TempRst1
        MsgBox "Inserted1"
    End If
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'    Temp_rst1
    cmdPrint.SetFocus
End Sub
Private Sub cmdShow_Click()
If cmdSave.Enabled = False Then Exit Sub

        If txtPat_ID1.text = "" Then
        MsgBox "Patient ID mandatory"
        Me.txtPat_ID = ""
        txtPat_ID1.SetFocus
        Exit Sub
    End If
'===for show data in Datagrid1=============
    Adodc1.connectionstring = strcn.Connection

    Adodc1.RecordSource = "exec Pro_FLUSH_TN 1,'" + txtM_Code + "','" + txtPat_ID + "'"
    Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.Visible = True
        DataGrid1.Columns(2).Width = 5250
        DataGrid1.Columns(0).Caption = "Group Code"
        DataGrid1.Columns(1).Caption = "Test Code"
        DataGrid1.Columns(2).Caption = "   Name of Test"
    Else
        DataGrid1.Visible = False
        MsgBox "Invalid Patient ID"
        txtPat_ID = ""
        txtPat_ID1 = ""
        txtPat_ID1.SetFocus
        Exit Sub
    End If
'===============================================
End Sub
Private Sub ComResult1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub
Private Sub ComResult2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub
Private Sub ComTest_Name_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub
Private Sub ComTest_Title_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub

Private Sub ComTest_Title_LostFocus()
    On Error Resume Next
    If ComTest_Title.text = "" Then
        cmdSave.SetFocus
        Exit Sub
    End If
    
    'If ComTest_Title.Text = "NOGU" Then
    'txtUsed_tech = " ( Not Applicable )"
    'Else
    'txtUsed_tech = ""
    'End If
    
    'GetResult
    GetS_Code
    GetUsed_Tech
End Sub

Private Sub DataGrid2_Click()
On Error Resume Next
    ComTest_Title = DataGrid2.Columns(0).value
    txtTest_Name = DataGrid2.Columns(0).value
End Sub

Private Sub Form_Click()
    DataGrid1.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    If rMicrobiology.DataGrid1.Visible = True Then
        rMicrobiology.DataGrid1.Visible = False
    Else
        Unload Me
    End If
    End If
End Sub

Private Sub Form_Load()
    Adodc8.connectionstring = strcn.Connection
    Adodc8.RecordSource = "exec m_name_select 2,'" + "MICROBIOLOGY" + "'"
    Adodc8.Refresh
    If Adodc8.Recordset.RecordCount > 0 Then
        txtM_Code = Adodc8.Recordset!m_code
    Else
        MsgBox "Inserted incurrect head name, first you have to insert currect name from TEST INFORMATION form then open this screen again"
        txtPat_ID.Enabled = False
        cmdSave.Enabled = False
    End If


    Dt.value = Now
    
'    GetTestName

    Temp_rst1
        
    StrScreenName = "Microbiology"
    Flush_Font_Type
    txtN_Exam = "As Under"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Temp_Table1 = Nothing

End Sub
Private Sub txtN_Exam_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub

Private Sub txtOrg_Iso_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim NValue As String
        Dim Power As String
        
        NValue = InputBox("Enter the value of n", "Power of n")

        For i = 0 To (Len(NValue) - 1)
            Select Case Mid(NValue, i + 1, 1)
            Case "0"
            Power = Power & PowerOfNone
            Case "1"
            Power = Power & PowerOfOne
            Case "2"
            Power = Power & PowerOfTwo
            Case "3"
            Power = Power & PowerOfThree
            Case "4"
            Power = Power & PowerOfFour
            Case "5"
            Power = Power & PowerOfFive
            Case "6"
            Power = Power & PowerOfSix
            Case "7"
            Power = Power & PowerOfSeven
            Case "8"
            Power = Power & PowerOfEight
            Case "9"
            Power = Power & PowerOfNine
            End Select
        Next i

        txtOrg_Iso.text = txtOrg_Iso.text & Power

End If

End Sub

Private Sub txtOrg_Iso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub

Private Sub txtPat_ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub
Private Sub DataGrid1_DblClick()
    'txtM_Code.Text = DataGrid1.Columns(0)
    txtS_Code.text = DataGrid1.Columns(1)
    txtS_Name.text = DataGrid1.Columns(2)
    txtSpecimen.SetFocus
    DataGrid1.Visible = False
End Sub

Private Sub txtPat_ID1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtPat_ID1_LostFocus()

If txtPat_ID1 = "" Then Exit Sub
       
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
    
   
    If IntPat_ID = 0 Then
        MsgBox "Invalid ID, Try again"
        txtPat_ID = ""
        txtPat_ID1 = ""
        'txtDummy_Pat_ID = ""
        txtPat_ID1.SetFocus
        Exit Sub
    End If

'----------------------------


If Len(Trim(txtPat_ID.text)) = 0 Then Exit Sub
    
Adodc5.connectionstring = strcn.Connection
Adodc5.RecordSource = "exec Pro_FLUSH 6," & txtPat_ID & ""
Adodc5.Refresh
If Adodc5.Recordset.RecordCount > 0 Then
    
Else
    '===for show data in Datagrid1=============
                Adodc1.connectionstring = strcn.Connection
                Adodc1.RecordSource = "exec Pro_FLUSH_TN 1,'" + txtM_Code + "','" + txtPat_ID + "'"
                Adodc1.Refresh
                If Adodc1.Recordset.RecordCount > 0 Then
                    'DataGrid1.Visible = True
                    DataGrid1.Columns(2).Width = 5270
                    DataGrid1.Columns(0).Caption = "Group Code"
                    DataGrid1.Columns(1).Caption = "Test Code"
                    DataGrid1.Columns(2).Caption = "   Name of Test"
                Else
                    DataGrid1.Visible = False
                    MsgBox "Invalied Patient ID"
                    txtPat_ID = ""
                    txtPat_ID1.SetFocus
                    
                End If
        '===============================================
End If

'--------------------------------------
'If Len(txtS_Code.Text) = 0 Then Exit Sub
If cmdSave.Enabled = False Then Exit Sub

If Trim(txtPat_ID1) = "" Then
    MsgBox "Patient ID mandatory"
    txtPat_ID1.SetFocus
    Exit Sub
End If

'If Len(Trim(txtS_Code)) = 0 Then Exit Sub
         
    Adodc6.connectionstring = strcn.Connection
    Adodc6.RecordSource = "exec Report_All_Select2 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
    Adodc6.Refresh
    
    If Adodc6.Recordset.RecordCount > 0 Then
        DataGrid1.Visible = False
        txtSpecimen = Adodc6.Recordset!Field1
        txtN_Exam = Adodc6.Recordset!Field2
        txtOrg_Iso.text = Adodc6.Recordset!Field3
        txtUsed_tech = Adodc6.Recordset!Field4
        ComTest_Title.text = Adodc6.Recordset!Field5
        'txtTest_Name.Text = Adodc6.Recordset!Field6
        txtSN.text = Adodc6.Recordset!Field14
        txtNote = Adodc6.Recordset!Field15
         Dt.value = Adodc6.Recordset!Dt
'

'/////////show in Temp_rst1//////////////
        con.connectionstring = strcn.Connection
        con.Open
        Temp_Table_Helper1.Open "select * from report_all where pat_id='" + txtPat_ID + "' and s_code='" + txtS_Code + "'and m_code='" + txtM_Code + "'", con

          While Temp_Table_Helper1.EOF = False
                Temp_Table1.AddNew
                Temp_Table1!Test_Name = Temp_Table_Helper1!Field6
                Temp_Table1!Test_result = Temp_Table_Helper1!Field7
                'Temp_Table1!unit = Temp_Table_Helper1!Field6
'
                Temp_Table_Helper1.MoveNext
            Wend
        DataGrid2.Refresh
        Temp_Table_Helper1.Close
        con.Close
'/////////end show in Temp_rst1////////////////////////////

    Else
    '===for show data in Datagrid1=============
                Adodc1.connectionstring = strcn.Connection
                Adodc1.RecordSource = "exec Pro_FLUSH_TN 1,'" + txtM_Code + "','" + txtPat_ID + "'"
                Adodc1.Refresh

                If Adodc1.Recordset.RecordCount > 0 Then
                    'DataGrid1.Visible = True
                    DataGrid1.Columns(2).Width = 5270
                    DataGrid1.Columns(0).Caption = "Group Code"
                    DataGrid1.Columns(1).Caption = "Test Code"
                    DataGrid1.Columns(2).Caption = "   Name of Test"
                Else
                    DataGrid1.Visible = False
                End If
'===============================================
    End If
    
Call BindTestName

End Sub

Private Sub txtS_Code_Change()
   '--------FOR S_NAME-------------------------
        Adodc8.connectionstring = strcn.Connection
        Adodc8.RecordSource = "exec S_name_select 1,'" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
        Adodc8.Refresh
        If Adodc8.Recordset.RecordCount > 0 Then
            txtS_Name.ForeColor = vbBlack
            txtS_Name = Adodc8.Recordset!s_name
        Else
            txtS_Name.ForeColor = vbRed
            txtS_Name = " Invalied Test code"
        End If
'--------END---------------------------------
End Sub
Private Sub txtS_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub
Private Sub txtS_Code_LostFocus()

'If Len(txtS_Code.Text) = 0 Then Exit Sub
If cmdSave.Enabled = False Then Exit Sub

If Trim(txtPat_ID1) = "" Then
    MsgBox "Patient ID mandatory"
    txtPat_ID1.SetFocus
    Exit Sub
End If

'If Len(Trim(txtS_Code)) = 0 Then Exit Sub
         
    Adodc6.connectionstring = strcn.Connection
    Adodc6.RecordSource = "exec Report_All_Select2 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
    Adodc6.Refresh
    
    If Adodc6.Recordset.RecordCount > 0 Then
        DataGrid1.Visible = False
        txtSpecimen = Adodc6.Recordset!Field1
        txtN_Exam = Adodc6.Recordset!Field2
        txtOrg_Iso.text = Adodc6.Recordset!Field3
        txtUsed_tech = Adodc6.Recordset!Field4
        ComTest_Title.text = Adodc6.Recordset!Field5
        txtTest_Name.text = Adodc6.Recordset!Field6
        txtTest_Result.text = Adodc6.Recordset!Field7
        txtSN.text = Adodc6.Recordset!Field14
        txtNote = Adodc6.Recordset!Field15
         Dt.value = Adodc6.Recordset!Dt
'
'/////////end show in Temp_rst1////////////////////////////

    Else
    '===for show data in Datagrid1=============
                Adodc1.connectionstring = strcn.Connection
                Adodc1.RecordSource = "exec Pro_FLUSH_TN 1,'" + txtM_Code + "','" + txtPat_ID + "'"
                Adodc1.Refresh

                If Adodc1.Recordset.RecordCount > 0 Then
                    DataGrid1.Visible = True
                    DataGrid1.Columns(2).Width = 5270
                    DataGrid1.Columns(0).Caption = "Group Code"
                    DataGrid1.Columns(1).Caption = "Test Code"
                    DataGrid1.Columns(2).Caption = "   Name of Test"
                Else
                    DataGrid1.Visible = False
                End If
'===============================================
    End If
    

End Sub


Private Sub InsReport_All_TempRst1()
    
    Temp_Table1.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    While Temp_Table1.EOF = False
    
          cmd.CommandText = "exec pro_Report_All 'I','" + Trim(txtPat_ID) + _
            "','" + txtM_Code + _
            "','" + txtS_Code + _
            "','" + ChkForQuote(txtSpecimen) + _
            "','" + ChkForQuote(txtN_Exam) + _
            "','" + ChkForQuote(txtOrg_Iso.text) + _
            "','" + ChkForQuote(txtUsed_tech) + _
            "','" + ChkForQuote(Trim(ComTest_Title)) + _
            "','" + ChkForQuote(Temp_Table1!Test_Name) + _
            "','" + ChkForQuote(Temp_Table1!Test_result) + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + ChkForQuote(Trim(txtSN.text)) + _
            "','" + ChkForQuote(txtNote) + _
            "','" + u_id + _
            "','" + Format(Dt, "yyyy-mm-dd") + _
            "','" + "" + _
            "','" + txtPat_ID1 + "'"
            cmd.Execute
            Temp_Table1.MoveNext
    Wend
    con.Close
End Sub

Private Sub Del_All_Report_All_TempRst1()
   
    'Temp_Table1.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    'While Temp_Table1.EOF = False
            cmd.CommandText = "exec Report_All_Delete2 1,'" + Trim(txtPat_ID.text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
            cmd.Execute
    '        Temp_Table1.MoveNext
    'Wend
    con.Close
End Sub
Private Sub txtSpecimen_dt_Time_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys Chr(9)
    End If
End Sub
Private Sub txtSpecimen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub



Private Sub txtUsed_tech_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub
Private Sub Clearscreen()
    'txtPat_ID = ""
    'txtM_Code = ""
    txtPat_ID1 = ""
    txtPat_ID.text = ""
    txtTest_Name.text = ""
    txtS_Code = ""
    txtS_Name = ""
    txtSpecimen.text = ""
    txtUsed_tech = ""
'    txtN_Exam = ""
    txtNote = ""
    txtTest_Result = ""
    ComTest_Title.Clear
    Dt.value = Date
    
End Sub
Private Sub GetTestName()
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select1 1,'" + "07" + "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
       Do Until Adodc7.Recordset.EOF
          ComTest_Title.AddItem Adodc7.Recordset!Test_Name
       Adodc7.Recordset.MoveNext
       Loop
    End If
End Sub
Private Sub GetResult()
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select 1,'" + ComTest_Title + "','" + "07" + "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
        txtSpecimen.text = Adodc7.Recordset!Test_result
        txtTest_Name.text = Adodc7.Recordset!unit
        txtTest_Result.text = Adodc7.Recordset!ref_range
        'txtUnit = Adodc7.Recordset!unit
        'txtRef_Range = Adodc7.Recordset!ref_range
    End If
End Sub
Private Sub Pat_Paid()
    Adodc5.connectionstring = strcn.Connection
    Adodc5.RecordSource = "exec Select_Paid 1,'" + txtPat_ID + "'"
    Adodc5.Refresh
    If Adodc5.Recordset.RecordCount > 0 Then
    Else
        txtPat_ID = ""
        MsgBox " Patient could not paid"
    End If
    
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


Private Sub BindTestName()
    On Error GoTo err_loop
    ComTest_Title.Clear
       con.connectionstring = strcn.Connection
       con.Open
       rs.Open "exec GetTestName '" & Trim(txtM_Code.text) & "','" & Trim(txtPat_ID.text) & "'", con

       If rs.EOF = False Then
          Do Until rs.EOF
            ComTest_Title.AddItem rs!Test_Name
          rs.MoveNext
          Loop
       End If
       rs.Close
       con.Close
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub GetS_Code()

  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_Result_Select8 '" & txtPat_ID & "','" & ComTest_Title.text & "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    Do Until Adodc7.Recordset.EOF
    
    
    txtTest_Name.text = Adodc7.Recordset!unit
    'txtTest_Result = Adodc7.Recordset!ref_range
    
    Adodc7.Recordset.MoveNext
    Loop

    End If
End Sub

Private Sub GetUsed_Tech()

  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_Result_Select12 '" & txtPat_ID.text & "','" & ComTest_Title.text & "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    Do Until Adodc7.Recordset.EOF
    '
    'ComTest_Name.AddItem Adodc7.Recordset!test_result
    'txtTest_Result = Adodc7.Recordset!test_result
    'txtUnit.Text = Adodc7.Recordset!unit
    txtSpecimen.text = Adodc7.Recordset!Test_result
        
    txtOrg_Iso.text = Adodc7.Recordset!others
    txtUsed_tech = Adodc7.Recordset!others1
    
    Adodc7.Recordset.MoveNext
    Loop

    End If
End Sub

Private Sub Del_Report()

    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
            cmd.CommandText = "exec Del_Report 1,'" + Trim(txtPat_ID.text) + "'"
            cmd.Execute
    con.Close
End Sub

Public Sub Temp_rst1()
    
    Set Temp_Table1 = New ADODB.Recordset
    With Temp_Table1
        .Fields.Append "Test_Name", adVarChar, 500
        .Fields.Append "Test_Result", adVarChar, 500
        '.Fields.Append "Unit", adVarChar, 500
        '.Fields.Append "Ref_Range", adVarChar, 500
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set DataGrid2.DataSource = Temp_Table1
    
    DataGrid2.Columns(0).DataField = "Test_Name"
    DataGrid2.Columns(1).DataField = "Test_Result"
    'DataGrid2.Columns(2).DataField = "Unit"
    'DataGrid2.Columns(3).DataField = "Ref_Range"
    DataGrid2.ReBind
    DataGrid2.Refresh
    
    DataGrid2.Columns(0).Width = 2800.142
    DataGrid2.Columns(1).Width = 5000
    'DataGrid2.Columns(2).Width = 3000
    'DataGrid2.Columns(3).Width = 0


End Sub
Private Sub DelReport_All_TempRst1()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
            cmd.CommandText = "exec Report_All_Delete 1,'" + Trim(txtPat_ID.text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + Trim(ComTest_Title) + "','" + "" + "'"
            cmd.Execute
    con.Close
End Sub
