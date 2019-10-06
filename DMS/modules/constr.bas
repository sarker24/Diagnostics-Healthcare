Attribute VB_Name = "connectionstring"
Option Explicit
'Public Count_Qty As Integer
Public con    As New ADODB.Connection
Public cmd    As New ADODB.Command
Public RS     As New ADODB.Recordset
Public strcn        As New Class1 '(Connection)
Public uid As String 'now not using
Public u_id As String 'for security option
Public upass As String 'for security option
Public Tflag As Boolean
Public pick As String
Public opr As String
Public StPat_ID As String 'show PAT_ID from PAT_INFO_MAIN
Public CRViewer1_MODE As String
Public Doc_List_MODE As String
Public Test_List_Mode As String
Public User_List_Mode As String
Public Item_List_MODE As String
Public Emp_List_MODE As String
Public Booth As String
Public BoothN As String
Public strScr_Name As String
Public strAllow As String
Public Source_Path As String
Public Destin_Path As String
Public StrRef_Code_N As String
Public IntFont As Integer
Public StrScreenName As String
Public NdocMode As String
Public StrItem_Code As String
Public StrEmp_Code As String
Public StrSub_Code As String
'Public StrDT As String
'Public StrMonth As String
'Public StrMonth1 As String
Public StDoc_Name As String
Public StrPat_ID_R As String

Public Sub Main()
    frmMAIN.Show vbModal
    
    
End Sub
Public Sub Locate_Booth()


Dim Boothdata As String
Dim FileNumber As Integer
FileNumber = FreeFile
Open App.Path + "\booth.dat" For Input Access Read As #FileNumber
Input #FileNumber, Boothdata
Close #FileNumber

If Boothdata <> "" Then
    Booth = Boothdata
    BoothN = Booth
    
Else
    MsgBox "Set Booth Name in Booth"
    Exit Sub
    End
End If
End Sub

Public Function pad(padtype As String, maxlength As Integer, value As Integer, padwith As String)
Dim i As Integer
Dim padvalue As String

    For i = 1 To maxlength - Len(CStr(value))
        padvalue = padvalue + padwith
    Next
    If UCase(padtype) = "L" Then
        pad = padvalue + Trim(CStr(value))
    End If
    If UCase(padtype) = "R" Then
        pad = Trim(CStr(value)) + padvalue
    End If

End Function
Public Function ChkForQuote(str As String)
Dim slposl As Integer, ln As Integer
Dim str2 As String
str2 = ""
For ln = 1 To Len(str)
    If Mid$(str, ln, 1) = "'" Then
        str2 = str2 + "'" + Mid(str, ln, 1)
    Else
        str2 = str2 + Mid$(str, ln, 1)
    End If
Next
ChkForQuote = str2
End Function
Public Sub Authority()
    strAllow = 0
    
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    'My_Rst.Open "select u_id from micropass", con
    My_Rst.Open "exec Select_Soft_Sucurity1 1,'" + u_id + "','" + strScr_Name + "'", con
    If My_Rst.EOF = False Then
    strAllow = My_Rst!allow
    End If
    con.Close
End Sub
Public Sub Flush_Font_Type()
    
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    'My_Rst.Open "select u_id from micropass", con
    My_Rst.Open "exec pro_name_SELECT '18','" & StrScreenName & "'", con
    If My_Rst.EOF = False Then
        IntFont = My_Rst!font_type
    End If
    con.Close
End Sub

Public Sub Source_Destin()
    
    Dim My_Rst As New ADODB.Recordset
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    My_Rst.Open "select * from Source_Destin", con
    If My_Rst.EOF = False Then
    Source_Path = My_Rst.Fields(0)
    Destin_Path = My_Rst.Fields(1)
    End If
    con.Close
End Sub

Public Sub Flush_Doc_Name()
    
    
    Dim My_R1 As New ADODB.Recordset
    
    con.connectionstring = strcn.Connection
    con.Open
    Set cmd.ActiveConnection = con
    
'    My_R1.Open "exec doc_name 1,'" + StrPat_ID_R + "'", con
    My_R1.Open "exec doc_name1 1,'" + StrPat_ID_R + "'", con
    If My_R1.EOF = False Then
'        StDoc_Name = My_R1!doc_name
    StDoc_Name = My_R1!cons
    End If
    con.Close
        
End Sub

Sub Gitna(frm As Form)
    frm.Left = (frmMAIN.ScaleWidth - frm.Width) / 2
    frm.Top = (frmMAIN.ScaleHeight - frm.Height) / 2
End Sub

'Public Sub Make_Month()
'StrMonth1 = ""
'StrDT = rDoc_Pay.stDt
'StrMonth = Mid(StrDT, 4, 2)
'
'Select Case StrMonth
'        Case "01"
'            StrMonth1 = "January"
'        Case "02"
'            StrMonth1 = "February"
'        Case "03"
'            StrMonth1 = "March"
'        Case "04"
'            StrMonth1 = "April"
'        Case "05"
'            StrMonth1 = "May"
'        Case "06"
'            StrMonth1 = "June"
'        Case "07"
'            StrMonth1 = "July"
'        Case "08"
'            StrMonth1 = "August"
'        Case "09"
'            StrMonth1 = "September"
'        Case "10"
'            StrMonth1 = "October"
'        Case "11"
'            StrMonth1 = "november"
'        Case "12"
'            StrMonth1 = "December"
'
'End Select
'MsgBox StrMonth1
'End Sub

