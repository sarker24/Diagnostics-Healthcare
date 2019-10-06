VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form lvItemSearch 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Item Search"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   Icon            =   "lvItemSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6075
   ScaleMode       =   0  'User
   ScaleWidth      =   10984.12
   StartUpPosition =   1  'CenterOwner
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
      Left            =   8040
      Picture         =   "lvItemSearch.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      Width           =   8055
   End
   Begin MSComctlLib.ListView lvSPD 
      Height          =   5655
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9975
      SortKey         =   1
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sub Code"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Main Code"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Test Name"
         Object.Width           =   8773
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Rate"
         Object.Width           =   1765
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Type"
         Object.Width           =   1765
      EndProperty
   End
End
Attribute VB_Name = "lvItemSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim str As String
Dim rs As New ADODB.Recordset

Private Sub cmdFind_Click()
con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con


     str = "select distinct a.m_code,b.s_code,b.s_name,c.rate,b.type,b.SerialNo from " & _
            "test_info_main a,test_info_sub b, test_info_rate c where " & _
            "a.m_code = b.m_code And a.m_code = c.m_code And b.s_code = c.s_code " & _
            " And b.s_name LIKE '" & txtSearch.text & "%'"

           If rs.State <> 0 Then rs.Close
              rs.Open str, con, adOpenStatic, adLockReadOnly

'    Form_Load
lvSPD.ListItems.Clear
    While Not rs.EOF

        With lvSPD.ListItems.Add

            .text = rs!SerialNo
            .SubItems(1) = rs!m_code
'            .SubItems(2) = Format(RS!PurchaseDate, "dd/MM/yyyy")
            .SubItems(2) = rs!s_name
            .SubItems(3) = rs!Rate
            .SubItems(4) = rs!Type
'            .SubItems(5) = RS!Type
        End With
        rs.MoveNext
    Wend

    con.Close
End Sub

Private Sub Form_Load()
'    con.connectionstring = strcn.Connection
'    con.Open
'    Set cmd.ActiveConnection = con
'    Call start_position(lvItemSearch)
    Call PopulateData
'    txtSearch.SetFocus
End Sub

Private Sub PopulateData()
'con.Close
 con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con


     str = "select distinct a.m_code,b.s_code,b.s_name,c.rate,b.type,b.SerialNo from " & _
            "test_info_main a,test_info_sub b, test_info_rate c where " & _
            "a.m_code = b.m_code And a.m_code = c.m_code And b.s_code = c.s_code " & _
            " And b.s_name LIKE '" & frmPatient_Info.txtSearch.text & "%'"

           If rs.State <> 0 Then rs.Close
              rs.Open str, con, adOpenStatic, adLockReadOnly



    While Not rs.EOF

        With lvSPD.ListItems.Add

'            .SubItems(1) = RS!m_code
'            .Text = RS!s_code
''            .SubItems(1) = RS!m_code
               .text = rs!SerialNo

'              .Text = RS!s_code
            .SubItems(1) = rs!m_code
            .SubItems(2) = rs!s_name
            .SubItems(3) = rs!Rate
            .SubItems(4) = rs!Type
'            .SubItems(5) = RS!Type
        End With
        rs.MoveNext
    Wend

    con.Close

End Sub



Private Sub lvSPD_DblClick()
 con.connectionstring = strcn.Connection
    con.Open

    Dim i As Integer
    If lvSPD.SelectedItem Is Nothing Then
        Unload Me
        Exit Sub
    End If

str = "select distinct a.m_code,b.s_code,b.s_name,c.rate,b.type  from " & _
            "test_info_main a,test_info_sub b, test_info_rate c where " & _
            "a.m_code = b.m_code And a.m_code = c.m_code And b.s_code = c.s_code " & _
            " And b.SerialNo= " & Val(lvSPD.SelectedItem.text) & " order by a.m_code,b.s_code"


    If rs.State <> 0 Then rs.Close
    rs.Open str, con, adOpenStatic, adLockReadOnly



                    frmPatient_Info.txtM_Code = rs!m_code
                    frmPatient_Info.txtS_Code = rs!s_code
                    frmPatient_Info.txtS_Name = rs!s_name
                    frmPatient_Info.nbrTest_Rate = rs!Rate
                    frmPatient_Info.txtType = rs!Type

                    Unload Me
                    frmPatient_Info.nbrTest_Rate.SetFocus
    Unload Me
    con.Close
End Sub



Private Sub lvSPD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
con.connectionstring = strcn.Connection
    con.Open

    Dim i As Integer
    If lvSPD.SelectedItem Is Nothing Then
        Unload Me
        Exit Sub
    End If



      str = "select distinct a.m_code,b.s_code,b.s_name,c.rate,b.type from " & _
            "test_info_main a,test_info_sub b, test_info_rate c where " & _
            "a.m_code = b.m_code And a.m_code = c.m_code And b.s_code = c.s_code " & _
            " And b.SerialNo= " & Val(lvSPD.SelectedItem.text) & " order by a.m_code,b.s_code"


    If rs.State <> 0 Then rs.Close
    rs.Open str, con, adOpenStatic, adLockReadOnly



                    frmPatient_Info.txtM_Code = rs!m_code
                    frmPatient_Info.txtS_Code = rs!s_code
                    frmPatient_Info.txtS_Name = rs!s_name
                    frmPatient_Info.nbrTest_Rate = rs!Rate
                    frmPatient_Info.txtType = rs!Type

                    Unload Me
'                    frmPatient_Info.nbrTest_Rate.SetFocus
                     frmPatient_Info.Delv_TM.SetFocus
    Unload Me
    con.Close
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



