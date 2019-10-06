VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form rHaematology 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Lab Report Format [HAEMATOLOGY]"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   DrawWidth       =   2
   Icon            =   "rHaematology.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtUsed_tech 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2130
      MultiLine       =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1290
      Width           =   7290
   End
   Begin VB.TextBox txtPat_ID1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   0
      ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
      Top             =   870
      Width           =   1260
   End
   Begin VB.TextBox txtGroup_name 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   450
      TabIndex        =   12
      Top             =   6990
      Width           =   2715
   End
   Begin VB.ComboBox ComTest_Title1 
      DataSource      =   "Adodc7"
      Height          =   315
      Left            =   420
      TabIndex        =   11
      Top             =   6630
      Width           =   10305
   End
   Begin VB.TextBox txtTest_Result1 
      BorderStyle     =   0  'None
      Height          =   2040
      Left            =   3180
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   6990
      Width           =   7560
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "rHaematology.frx":000C
      Height          =   780
      Left            =   2880
      TabIndex        =   25
      Top             =   1830
      Visible         =   0   'False
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   1376
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
   Begin VB.TextBox txtPat_ID 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2175
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   1
      Top             =   855
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txtM_Code 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3645
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   855
      Width           =   345
   End
   Begin VB.TextBox txtS_Code 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   5130
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox txtS_Name 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   6105
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   2145
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
      Height          =   300
      Left            =   9480
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   840
      Width           =   900
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   10230
      Width           =   1050
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFFFFF&
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
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   10230
      Width           =   1050
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
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
      Left            =   9780
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   10230
      Width           =   1050
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   10230
      Width           =   1050
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   10230
      Width           =   1050
   End
   Begin VB.TextBox txtNote 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   390
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   9180
      Width           =   10380
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
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
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   10230
      Width           =   1050
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   9960
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   7800
      Top             =   0
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Left            =   7770
      Top             =   0
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
      Left            =   8730
      Top             =   0
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1665
      Left            =   420
      TabIndex        =   24
      Top             =   4860
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   2937
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
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
         Name            =   "Arial"
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
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2640.189
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8730
      Top             =   -30
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
      Left            =   8730
      Top             =   0
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
      Left            =   9810
      Top             =   30
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Left            =   8730
      Top             =   -90
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
      Left            =   8730
      Top             =   30
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
      Left            =   8205
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      _Version        =   393216
      Format          =   65273857
      CurrentDate     =   37114
   End
   Begin VB.ComboBox ComTest_Title 
      DataSource      =   "Adodc7"
      Height          =   315
      Left            =   420
      TabIndex        =   7
      Top             =   1830
      Width           =   2415
   End
   Begin RichTextLib.RichTextBox ComTest_Name 
      Height          =   2415
      Left            =   2850
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2430
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4260
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"rHaematology.frx":0021
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
   Begin RichTextLib.RichTextBox txtTest_Result 
      Height          =   2415
      Left            =   5490
      TabIndex        =   9
      Top             =   2430
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4260
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"rHaematology.frx":009A
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
   Begin RichTextLib.RichTextBox txtNormal_Value 
      Height          =   2415
      Left            =   8130
      TabIndex        =   10
      Top             =   2430
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   4260
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"rHaematology.frx":0113
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
   Begin RichTextLib.RichTextBox txtType 
      Height          =   2415
      Left            =   10710
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2430
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   4260
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"rHaematology.frx":018C
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
   Begin VB.CommandButton cmdDelete_TempTable1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "D E L E T E"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2445
      Width           =   345
   End
   Begin VB.TextBox txtSN 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   390
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   9975
      Width           =   10380
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Used Technology"
      Height          =   195
      Left            =   450
      TabIndex        =   36
      Top             =   1230
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special Note"
      Height          =   195
      Left            =   390
      TabIndex        =   34
      Top             =   9690
      Width           =   915
   End
   Begin VB.Label lblOverflow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
      Height          =   195
      Left            =   5280
      TabIndex        =   33
      Top             =   1560
      Width           =   450
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Test"
      Height          =   195
      Left            =   450
      TabIndex        =   32
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   7650
      TabIndex        =   31
      Top             =   870
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID"
      Height          =   195
      Left            =   465
      TabIndex        =   30
      Top             =   870
      Width           =   705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Code"
      Height          =   195
      Left            =   5100
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   6120
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HAEMATOLOGICAL ANALYSIS REPORT"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   555
      TabIndex        =   27
      Top             =   285
      Width           =   3840
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Normal Value"
      Height          =   195
      Left            =   8880
      TabIndex        =   26
      Top             =   1560
      Width           =   945
   End
End
Attribute VB_Name = "rHaematology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp_Table1 As New ADODB.Recordset
Dim Temp_Table_Helper1 As New ADODB.Recordset
Dim Temp_Table2 As New ADODB.Recordset
Dim Temp_Table_Helper2 As New ADODB.Recordset

Dim Strpat_id1 As String
Dim StrRow_Count As String
Dim StrPat_Type As String
Dim IntPat_ID As Double

Private Sub cmdClear_Click()
        
    ComTest_Title.text = ""
    ComTest_Name.text = ""
    ComTest_Title1 = ""
    txtTest_Result1 = ""
    txtPat_ID.text = ""
    txtS_Code.text = ""
    txtS_Name.text = ""
    Dt.value = Now
    txtNote = ""
    txtSN.text = ""
    txtN_Exam = ""
    txtNormal_Value = ""
    txtTest_Result = ""
    txtGroup_name = ""
'    txtUnit = ""
    txtSpecimen = ""
    txtUsed_tech = ""
    
    Temp_rst1
    txtPat_ID = ""
    txtPat_ID1 = ""
    txtType.text = ""
    DataGrid1.Visible = False
    txtPat_ID1.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdDelete_Click()
    If cmdSave.Enabled = False Then Exit Sub

    Dim Strmsg As String
    Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
    If Strmsg = vbYes Then
    'Del_All_Report_All_TempRst1
    Del_Report
    Clearscreen
    
    txtPat_ID = ""
    txtPat_ID1 = ""
    GetTestTitle1
    
    txtPat_ID1.SetFocus
    End If
End Sub
Private Sub cmdDelete_TempTable1_Click()

    If ComTest_Name = "" Then Exit Sub
    If cmdSave.Enabled = False Then Exit Sub
    If Temp_Table1.RecordCount <= 0 Then Exit Sub
    
    If Trim(ComTest_Name.text) = "" Then
        MsgBox "You didn't select the the Test Name"
        DataGrid2.SetFocus
        Exit Sub
    Else
        Dim Strmsg As String
        Strmsg = MsgBox("Do you want to delete?", vbQuestion + vbYesNo)
        If Strmsg = vbYes Then
        DelReport_All_TempRst1
        Temp_Table1.Delete
        ComTest_Name = ""
        txtTest_Result = ""
'        txtUnit = ""
        txtNormal_Value = ""
        End If
        
    End If
End Sub

Private Sub CmdPreview_Click()
    CRViewer1_MODE = 28
    Viewer.Show vbModal
End Sub

Private Sub cmdPrint_Click()
    '==========direct print==========================
            
            Dim Report28 As New Haematology
            Dim StrPat_ID As String
           
            
            Dim strM_Code As String
            Dim strS_Code As String
            
            StrPat_ID = rHaematology.txtPat_ID
            StrPat_ID_R = StrPat_ID
            strM_Code = rHaematology.txtM_Code
            strS_Code = rHaematology.txtS_Code
            
             '--------------------------------------------------------------------
            Report28.FormulaFields.Item(1).text = Chr(34) & IntFont & Chr(34)
            Report28.FormulaFields.Item(2).text = Chr(34) & "Patient ID" & Chr(34)
            Report28.FormulaFields.Item(3).text = Chr(34) & "Received Date" & Chr(34)
            Report28.FormulaFields.Item(4).text = Chr(34) & "Delivered Date" & Chr(34)
            Report28.FormulaFields.Item(5).text = Chr(34) & "Patient Name" & Chr(34)
            Report28.FormulaFields.Item(6).text = Chr(34) & "Age" & Chr(34)
            Report28.FormulaFields.Item(7).text = Chr(34) & "Sex" & Chr(34)
            Report28.FormulaFields.Item(8).text = Chr(34) & "Refd. by" & Chr(34)
            '--------------------------------------------------------------------
            Report28.FormulaFields.Item(9).text = Chr(34) & "Haematological Report" & Chr(34)
            Report28.FormulaFields.Item(10).text = Chr(34) & "Tests" & Chr(34)
            Report28.FormulaFields.Item(11).text = Chr(34) & "Results" & Chr(34)
            Report28.FormulaFields.Item(12).text = Chr(34) & "Normal Values" & Chr(34)
            Report28.FormulaFields.Item(13).text = Chr(34) & "Checked By" & Chr(34)
            If rHaematology.txtS_Code = "23" Or rHaematology.txtS_Code = "24" Then
            Report28.FormulaFields.Item(14).text = Chr(34) & "Test are carried out by SYSMEX KX-21" & Chr(34)
            End If
            Report28.Text2.SetText Trim(rHaematology.txtTest_Result1.text)
            
            Call Flush_Doc_Name
            Report28.Text4.SetText StDoc_Name
            
            Report28.DiscardSavedData
            rs.Open "exec Rpt 1,'" + StrPat_ID + "','" + strM_Code + "','" + strS_Code + "'", strcn.Connection
            Report28.Database.SetDataSource rs
            
            Report28.PrintOut (False)
            rs.Close
            
            Call cmdClear_Click
            txtPat_ID1.SetFocus
    '====================================
    
End Sub

Private Sub cmdSave_Click()
'-----validation check---------------------
    If Trim(txtPat_ID) = "" Then
        MsgBox "Patient ID mandatory"
        txtPat_ID.SetFocus
        Exit Sub
    End If
    
'    If Len(Trim(txtS_Code)) = 0 Then
'        MsgBox "Test Code mandatory"
'        txtS_Code.SetFocus
'        Exit Sub
'    End If
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
    Temp_rst1
    'txtGroup_name = ""
    'txtTest_Result1 = ""
    'txtNote = ""
    txtUsed_tech.text = ""
    DataGrid1.Visible = False
    cmdPrint.SetFocus
    
End Sub
Private Sub cmdShow_Click()
If cmdSave.Enabled = False Then Exit Sub

        If txtPat_ID1.text = "" Then
        MsgBox "Patient ID mandatory"
        txtPat_ID = ""
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
        txtPat_ID1 = ""
        txtPat_ID = ""
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

Private Sub ComTest_Name_LostFocus()
'    GetResult
End Sub

Private Sub ComTest_Title_LostFocus()
'    Trim(ComTest_Name.Text) = ""
    'GetTestName
    If ComTest_Title = "" Then
        On Error Resume Next
        cmdSave.SetFocus
        Exit Sub
    End If
    
    GetS_Code
    GetUsed_Tech
End Sub
Private Sub ComTest_Title1_LostFocus()
    GetResult1
End Sub
Private Sub DataGrid2_DblClick()
On Error Resume Next
    ComTest_Title.text = DataGrid2.Columns(0)
    ComTest_Name.text = DataGrid2.Columns(1)
    txtTest_Result.text = DataGrid2.Columns(2)
    txtTest_Result.text = DataGrid2.Columns(2)
'    txtUnit.Text = DataGrid2.Columns(2)
    txtNormal_Value = DataGrid2.Columns(3)
    txtType.text = DataGrid2.Columns(4)
End Sub

Private Sub Form_Click()
    If DataGrid1.Visible = True Then
       DataGrid1.Visible = False
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    If rHaematology.DataGrid1.Visible = True Then
        rHaematology.DataGrid1.Visible = False
    Else
        Unload Me
    End If
    End If
End Sub

Private Sub Form_Load()
    
    
    Adodc8.connectionstring = strcn.Connection
    Adodc8.RecordSource = "exec m_name_select 2,'" + "HAEMATOLOGY" + "'"
    Adodc8.Refresh
    If Adodc8.Recordset.RecordCount > 0 Then
        txtM_Code = Adodc8.Recordset!m_code
    Else
        MsgBox "Inserted correct head name, first you have to insert currect name from TEST INFORMATION form then open this screen again"
        txtPat_ID.Enabled = False
        cmdSave.Enabled = False
    End If


    Dt.value = Now
    Temp_rst1
'------select data from test_result------
'     GetTestTitle
     
     
'     GetTestName
'------end------------------------------

DataGrid2.Columns(0).Width = 2429.858
DataGrid2.Columns(1).Width = 2640.189
'DataGrid2.Columns(2).Width = 1769.953
'DataGrid2.Columns(3).Width = 1785.26
StrScreenName = "Haematology"
Flush_Font_Type
    
GetTestTitle1

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Temp_Table1 = Nothing

End Sub
Private Sub txtN_Exam_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub

Private Sub txtNormal_Value_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub

Private Sub txtPat_ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     SendKeys Chr(9)
    End If
End Sub
Private Sub txtPat_ID_LostFocus()
'Pat_Paid
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
                    DataGrid1.Visible = True
                    DataGrid1.Columns(2).Width = 5270
                    DataGrid1.Columns(0).Caption = "Group Code"
                    DataGrid1.Columns(1).Caption = "Test Code"
                    DataGrid1.Columns(2).Caption = "   Name of Test"
                Else
                    DataGrid1.Visible = False
                    MsgBox "Invalied Patient ID"
                    txtPat_ID = ""
                    txtPat_ID.SetFocus
                    
                End If
        '===============================================
End If


End Sub
Private Sub DataGrid1_DblClick()
    'txtM_Code.Text = DataGrid1.Columns(0)
    txtS_Code.text = DataGrid1.Columns(1)
    StrSub_Code = txtS_Code.text
    txtS_Name.text = DataGrid1.Columns(2)
    ComTest_Name.SetFocus
    DataGrid1.Visible = False
    
End Sub
Private Sub txtNormal_Value_LostFocus()
If Trim(ComTest_Name) = "" Then Exit Sub
'----------------check--------
Dim Check As Integer
Check = 0
If Temp_Table1.RecordCount > 0 Then
    Temp_Table1.MoveFirst
    
        While Temp_Table1.EOF = False
                
            If Temp_Table1!Test_Name = ComTest_Title.text Then
            'And Temp_Table1!test_result = txtTest_Result.Text Then
                Check = 1
            End If
    Temp_Table1.MoveNext
        Wend
    If Check = 1 Then
        MsgBox "This Test Name already exists"
        Check = 0
        ComTest_Name = ""
        txtTest_Result = ""
'        txtUnit = ""
        txtNormal_Value = ""
        txtType.text = ""
        ComTest_Title.SetFocus
        Exit Sub
    End If
'    Temp_Table.MoveFirst
End If

'--------------end check-----

'+++to insert into TEMPORARY RECORDSET "Temp_rst1"++++
        Temp_Table1.AddNew
        Temp_Table1!Test_Name = ComTest_Title.text
        Temp_Table1!Test_result = ComTest_Name.text
        Temp_Table1!unit = txtTest_Result.text
        Temp_Table1!ref_range = txtNormal_Value.text
        Temp_Table1!Type = txtType.text
        DataGrid2.Refresh
'+++++++++++++++++++++++++++++++++++++++
'    DataGrid2.Columns(0).Width = 2000
'    DataGrid2.Columns(0).Width = 1000
    ComTest_Name = ""
    txtNormal_Value = ""
    txtTest_Result = ""
'    txtUnit = ""
    ComTest_Title = ""
    txtType.text = ""
    ComTest_Title.SetFocus
    
'ComTest_Name = ""
'txtTest_Result = ""
'txtUnit = ""
'txtNormal_Value = ""

'DataGrid2.Columns(0).Width = 2970.142
'DataGrid2.Columns(1).Width = 3300.095
'DataGrid2.Columns(2).Width = 1769.953
'DataGrid2.Columns(3).Width = 1785.26


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
'-->>>>>>>>>>>>

Temp_rst1
    
'If Len(txtS_Code.Text) = 0 Then Exit Sub
If cmdSave.Enabled = False Then Exit Sub

StrSub_Code = txtS_Code.text

If Trim(txtPat_ID1) = "" Then
    MsgBox "Patient ID mandatory"
    txtPat_ID1.SetFocus
    Exit Sub
End If


         
    Adodc6.connectionstring = strcn.Connection
    Adodc6.RecordSource = "exec Report_All_Select2 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
    Adodc6.Refresh
    
    If Adodc6.Recordset.RecordCount > 0 Then
        DataGrid1.Visible = False
'        txtSpecimen = Adodc6.Recordset!Field1
         ComTest_Title1.text = Adodc6.Recordset!Field6
         txtGroup_name.text = Adodc6.Recordset!Field7
         txtTest_Result1 = Adodc6.Recordset!Field8
         txtUsed_tech = Adodc6.Recordset!Field9
         Dt.value = Adodc6.Recordset!Dt

'++++++++++for show feild18 to txtNote +++++++++++
    Adodc8.connectionstring = strcn.Connection
    Adodc8.RecordSource = "exec Report_All_Select2 1,'" & Trim(txtPat_ID.text) & "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
    Adodc8.Refresh
    If Adodc8.Recordset.RecordCount > 0 Then
    txtSN.text = Adodc8.Recordset!Field14
    txtNote = Adodc8.Recordset!Field15
    End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
         
'/////////show in Temp_rst1//////////////
        con.connectionstring = strcn.Connection
        con.Open
        Temp_Table_Helper1.Open "select * from report_all where pat_id='" + txtPat_ID + "' and s_code='" + txtS_Code + "'and m_code='" + txtM_Code + "'", con
        
          While Temp_Table_Helper1.EOF = False
                Temp_Table1.AddNew
                Temp_Table1!Test_Name = Temp_Table_Helper1!Field1
                Temp_Table1!Test_result = Temp_Table_Helper1!Field2
                Temp_Table1!unit = Temp_Table_Helper1!Field3
                Temp_Table1!ref_range = Temp_Table_Helper1!Field4
                Temp_Table1!Type = Temp_Table_Helper1!Field5
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
    
DataGrid2.Columns(0).Width = 2429.858
DataGrid2.Columns(1).Width = 2640.189
DataGrid2.Columns(2).Width = 2640.189
DataGrid2.Columns(3).Width = 2600

'GetTestTitle

'GetTestTitle1



'--<<<<<<<<<<<<<<<

Call BindTestName
Call BindTestName1

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

GetTestTitle1

End Sub
Public Sub Temp_rst1()

    Set Temp_Table1 = New ADODB.Recordset
    With Temp_Table1
        .Fields.Append "Test_Name", adVarChar, 500
        .Fields.Append "Test_Result", adVarChar, 500
        .Fields.Append "Unit", adVarChar, 500
        .Fields.Append "Ref_Range", adVarChar, 500
        .Fields.Append "type", adVarChar, 1
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set DataGrid2.DataSource = Temp_Table1
    
    DataGrid2.Columns(0).DataField = "Test Goup"
    DataGrid2.Columns(1).DataField = "Test_Name"
    DataGrid2.Columns(2).DataField = "Test_Result"
    DataGrid2.Columns(3).DataField = "Ref_Range"
    DataGrid2.ReBind
    DataGrid2.Refresh
    DataGrid2.Columns(0).Caption = "Test Goup"
    DataGrid2.Columns(1).Caption = "Test_Name"
    DataGrid2.Columns(2).Caption = "Test_Result"
    DataGrid2.Columns(3).Caption = "Normal Values"
    
    DataGrid2.Columns(0).Width = 3284.788
    DataGrid2.Columns(1).Width = 2324.977
    DataGrid2.Columns(2).Width = 2310.236
    DataGrid2.Columns(3).Width = 2324.977
'End Sub
    
    
    '----------------
    
    
'End Sub

End Sub
Private Sub InsReport_All_TempRst1()
    On Error Resume Next
    Temp_Table1.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    While Temp_Table1.EOF = False
    
          cmd.CommandText = "exec pro_Report_All 'I','" + Trim(txtPat_ID) + _
            "','" + txtM_Code + _
            "','" + txtS_Code + _
            "','" + Temp_Table1!Test_Name + _
            "','" + Temp_Table1!Test_result + _
            "','" + Temp_Table1!unit + _
            "','" + Temp_Table1!ref_range + _
            "','" + Temp_Table1!Type + _
            "','" + Trim(ComTest_Title1) + _
            "','" + Trim(txtGroup_name.text) + _
            "','" + Trim(txtTest_Result1.text) + _
            "','" + Trim(txtUsed_tech.text) + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + "" + _
            "','" + Trim(txtSN.text) + _
            "','" + txtNote + _
            "','" + u_id + _
            "','" + Format(Dt, "yyyy-mm-dd") + _
            "','" + "" + _
            "','" + txtPat_ID1 + "'"
            cmd.Execute
            Temp_Table1.MoveNext
    Wend
    con.Close
End Sub

Private Sub DelReport_All_TempRst1()
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
            cmd.CommandText = "exec Report_All_Delete 1,'" + Trim(txtPat_ID.text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "','" + Trim(ComTest_Name) + "','" + "" + "'"
            cmd.Execute
    con.Close
End Sub
Private Sub Del_All_Report_All_TempRst1()
   On Error Resume Next
    Temp_Table1.MoveFirst
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    While Temp_Table1.EOF = False
            cmd.CommandText = "exec Report_All_Delete2 1,'" + Trim(txtPat_ID.text) + "','" + Trim(txtM_Code) + "','" + Trim(txtS_Code) + "'"
            cmd.Execute
            Temp_Table1.MoveNext
    Wend
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


Private Sub txtTest_Result_KeyPress(KeyAscii As Integer)
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
    txtS_Code = ""
    txtS_Name = ""
    txtSpecimen = ""
    txtUsed_tech = ""
    txtN_Exam = ""
    txtNote = ""
    txtType.text = ""
    ComTest_Title1.Clear
    txtGroup_name = ""
    txtTest_Result1 = ""
    Temp_rst1
    DataGrid1.Visible = False
    Dt.value = Date
    
End Sub


Private Sub GetTestTitle()

ComTest_Title.Clear

  Adodc7.connectionstring = strcn.Connection
  'Adodc7.RecordSource = "exec test_result_select1 1,'" + "01" + "'"
  Adodc7.RecordSource = "exec Flush_Test_Result 1,'" & txtM_Code & "','" & txtM_Code & "','" & StrSub_Code & "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
       Do Until Adodc7.Recordset.EOF
          ComTest_Title.AddItem Adodc7.Recordset!Test_Name
          'ComTest_Title = Adodc7.Recordset!Test_Name
       Adodc7.Recordset.MoveNext
       Loop
    End If
End Sub
Private Sub GetTestTitle1()
    ComTest_Title1.Clear
    
  Adodc7.connectionstring = strcn.Connection
  'Adodc7.RecordSource = "exec test_result_select1 1,'" + "01A" + "'"
  Adodc7.RecordSource = "exec Flush_Test_Result 2,'" & "01A" & "','" & txtM_Code & "',''"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
       Do Until Adodc7.Recordset.EOF
          ComTest_Title1.AddItem Adodc7.Recordset!Test_Name
       Adodc7.Recordset.MoveNext
       Loop
    End If
End Sub
Private Sub GetResult()
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select 2,'" + ComTest_Name.text + "','" + "01" + "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    txtTest_Result = Adodc7.Recordset!unit
    txtNormal_Value = Adodc7.Recordset!ref_range   'ref_range
    End If
End Sub
Private Sub GetResult1()
  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_result_select 1,'" + ComTest_Title1.text + "','" + "01A" + "'"
  Adodc7.Refresh
    If Adodc7.Recordset.RecordCount > 0 Then
    txtGroup_name.text = Adodc7.Recordset!Test_result
    txtTest_Result1.text = Adodc7.Recordset!unit
    txtNote = Adodc7.Recordset!ref_range
    End If
End Sub
Private Sub GetTestName()

  Adodc7.connectionstring = strcn.Connection
  'Adodc7.RecordSource = "exec test_result_select 1,'" + ComTest_Title + "','" + "01" + "'"
  Adodc7.RecordSource = "exec test_Result_Select7 1,'" & ComTest_Title.text & "','" & txtM_Code & "','" & StrSub_Code & "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    Do Until Adodc7.Recordset.EOF
    
    'ComTest_Name.AddItem Adodc7.Recordset!test_result
    ComTest_Name = Adodc7.Recordset!Test_result
    txtTest_Result = Adodc7.Recordset!unit
    txtNormal_Value = Adodc7.Recordset!ref_range
    
    Adodc7.Recordset.MoveNext
    Loop
    'txtUnit = Adodc7.Recordset!unit
    'txtUnit = Adodc7.Recordset!unit
    'txtNormal_Value = Adodc7.Recordset!ref_range   'ref_range
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
Private Sub BindTestName1()
'    On Error GoTo err_loop
'    ComTest_Title1.Clear
'       con.connectionstring = strcn.Connection
'       con.Open
'       RS.Open "exec GetTestName '" & "01A" & "','" & Trim(txtPat_ID.Text) & "'", con
'
'       If RS.EOF = False Then
'          Do Until RS.EOF
'            ComTest_Title1.AddItem RS!Test_Name
'          RS.MoveNext
'          Loop
'       End If
'       RS.Close
'       con.Close
'    Exit Sub
'err_loop:
'    MsgBox Err.Description, vbCritical
'    Exit Sub
End Sub

Private Sub GetS_Code()

  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_Result_Select8 '" & txtPat_ID & "','" & ComTest_Title.text & "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    Do Until Adodc7.Recordset.EOF
    
    'ComTest_Name.AddItem Adodc7.Recordset!test_result
    ComTest_Name = Adodc7.Recordset!Test_result
    txtTest_Result = Adodc7.Recordset!unit
    txtNormal_Value = Adodc7.Recordset!ref_range
    txtType.text = Adodc7.Recordset!others
    Adodc7.Recordset.MoveNext
    Loop

    End If
End Sub

Private Sub GetUsed_Tech()

  Adodc7.connectionstring = strcn.Connection
  Adodc7.RecordSource = "exec test_Result_Select15 '" & txtPat_ID & "','" & ComTest_Name.text & "'"
  Adodc7.Refresh

    If Adodc7.Recordset.RecordCount > 0 Then
    Do Until Adodc7.Recordset.EOF
    
    'ComTest_Name.AddItem Adodc7.Recordset!test_result
    'txtTest_Result = Adodc7.Recordset!test_result
    'txtUnit.Text = Adodc7.Recordset!unit
    txtUsed_tech.text = Adodc7.Recordset!others1
    
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

