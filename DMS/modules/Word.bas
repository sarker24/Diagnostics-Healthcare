Attribute VB_Name = "Word"
Option Explicit

Public Function ConvertX(nbrNumber As Double) As String
    Dim intDECI As Integer
    Dim intVAL As Integer
    Dim strNUM As String
    Dim strDECI As String
    Dim strHole As String
    Dim strWord As String
    
    intDECI = DecimalX(CStr(nbrNumber))
    
    If intDECI > 0 Then
       strNUM = Left$(nbrNumber, intDECI - 1)
       
        strDECI = Mid$(nbrNumber, intDECI + 1, 2)
        Dim L As String
        Dim R As String
        Dim strFlot As String
     '   Dim strFlot1 As String
        
        L = Left$(strDECI, 1)
        R = Mid$(strDECI, 2, 1)
             
        If L = "0" And R <> "0" Then
           strFlot = "  and zero " & Word1(R)
           GoTo stp
        End If
        
        If L <> "0" And R = "0" Then
           strFlot = " and " & Word1(L) & " zero "
           GoTo stp
        End If
        
'        If L <> "0" And R <> "0" Then
'           strFlot = " and " & Word1(L) & "  " & Word1(R)
'           GoTo stp
'        End If
        
        If L <> "" And R = "" Then
           strFlot = " and " & Word1(L) & " zero "
'           strFlot = strFlot1 & " zero "
        End If
stp:
    Else
       strNUM = nbrNumber
    End If
    
    intVAL = Len(strNUM)
    
    Select Case intVAL
        Case 1
            strHole = Word1(strNUM)
        Case 2
            strHole = Word2(strNUM)
        Case 3
            strHole = Word3(strNUM)
        Case 4
            strHole = Word4(strNUM)
        Case 5
            strHole = Word5(strNUM)
        Case 6
            strHole = Word6(strNUM)
        Case 7
            strHole = Word7(strNUM)
        Case 8
            strHole = Word8(strNUM)
        Case 9
            strHole = Word9(strNUM)
        Case 10
            strHole = Word10(strNUM)
        Case Else
            MsgBox "Supports 10 digit with two decimal places", vbInformation
            Exit Function
    End Select
    
    If strFlot = "" Then
       ConvertX = "Taka : (" & strHole & " only)"
    Else
       ConvertX = "Taka : (" & strHole & strFlot & " paisa only)"
    End If
    
    'MsgBox " " & nbrNumber.Text & " " & " In words( " & " " & strWord & " )"
End Function
Public Function Word1(intNum1 As String) As String
    If intNum1 <> "0" Then
       Select Case intNum1
       Case "1"
            Word1 = "One"
       Case "2"
            Word1 = "Two"
       Case "3"
            Word1 = "Three"
       Case "4"
            Word1 = "Four"
       Case "5"
            Word1 = "Five"
       Case "6"
            Word1 = "Six"
       Case "7"
            Word1 = "Seven"
       Case "8"
            Word1 = "Eight"
       Case "9"
            Word1 = "Nine"
       End Select
    Else
        Word1 = ""
    End If
End Function
Public Function Word2(intNum2 As String) As String
    Dim F As String
    Dim L As String
    
    F = Left$(intNum2, 1)
    L = Right$(intNum2, 1)
    
    If F = "0" Then
       Word2 = Word1(L)
       Exit Function
    Else
       If Val(F) >= 2 Then
          Select Case F
          Case "2"
            Word2 = "Twenty  " & Word1(L)
          Case "3"
            Word2 = "Thirty  " & Word1(L)
          Case "4"
            Word2 = "Fourty  " & Word1(L)
          Case "5"
            Word2 = "Fifty  " & Word1(L)
          Case "6"
            Word2 = "Sixty  " & Word1(L)
          Case "7"
            Word2 = "Seventy  " & Word1(L)
          Case "8"
            Word2 = "Eighty  " & Word1(L)
          Case "9"
            Word2 = "Ninty  " & Word1(L)
          End Select
       Else
          Select Case intNum2
          Case "10"
            Word2 = "Ten"
          Case "11"
            Word2 = "Eleven"
          Case "12"
            Word2 = "Twelve"
          Case "13"
            Word2 = "Thirteen"
          Case "14"
            Word2 = "Fourteen"
          Case "15"
            Word2 = "Fifteen"
          Case "16"
            Word2 = "Sixteen"
          Case "17"
            Word2 = "Seventeen"
          Case "18"
            Word2 = "Eightteen"
          Case "19"
            Word2 = "Ninghteen"
          End Select
       End If
    End If
End Function

Public Function Word3(intNUM3 As String) As String
    Dim F As String
    Dim L As String
    
    F = Left$(intNUM3, 1)
    L = Right$(intNUM3, 2)
    
    If F = "0" Then
        Word3 = Word2(L)
    Else
        Word3 = Word1(F) & " Hundred " & Word2(L)
    End If
End Function

Public Function Word4(intNUM4 As String) As String
    Dim F As String
    Dim L As String
    
    F = Left$(intNUM4, 1)
    L = Right$(intNUM4, 3)
    
    If F = "0" Then
        Word4 = Word3(L)
    Else
        Word4 = Word1(F) & " Thousand " & Word3(L)
    End If
End Function
Public Function Word5(intNUM5 As String) As String
    Dim LF As String
    Dim RT As String
    Dim LS As String
    Dim RF As String
    
    LF = Left$(intNUM5, 1)
    LS = Left$(intNUM5, 2)
    RT = Right$(intNUM5, 3)
    RF = Right$(intNUM5, 4)
    
    If LF = "0" Then
        Word5 = Word4(RF)
    Else
        Word5 = Word2(LS) & " Thousand " & Word3(RT)
    End If
End Function

Public Function Word6(intNUM6 As String) As String
    Dim LF As String
    Dim RT As String
    
    LF = Left$(intNUM6, 1)
    RT = Right$(intNUM6, 5)
    
    If LF = "0" Then
        Word6 = Word5(RT)
    Else
        Word6 = Word1(LF) & " Lac(s) " & Word5(RT)
    End If
End Function
Public Function Word7(intNUM7 As String) As String
    Dim LF As String
    Dim LS As String
    Dim RF As String
    Dim RS As String
    
    LF = Left$(intNUM7, 1)
    LS = Left$(intNUM7, 2)
    RF = Right$(intNUM7, 5)
    RS = Right$(intNUM7, 6)
    
    If LF = "0" Then
        Word7 = Word6(RS)
    Else
        Word7 = Word2(LS) & " Lacs " & Word5(RF)
    End If
End Function
Public Function Word8(intNUM8 As String) As String
    Dim LF As String
    Dim LS As String
    Dim RF As String
    Dim RS As String
    
    LF = Left$(intNUM8, 1)
    RS = Right$(intNUM8, 7)
    
    If LF = "0" Then
        Word8 = Word7(RS)
    Else
        Word8 = Word1(LF) & " Crore " & Word7(RS)
    End If
End Function
Public Function Word9(intNUM9 As String) As String
    Dim LF As String
    Dim LS As String
    Dim RF As String
    Dim RS As String
    
    LF = Left$(intNUM9, 1)
    LS = Left$(intNUM9, 2)
    RS = Right$(intNUM9, 7)
    RF = Right$(intNUM9, 8)
    
    If LF = "0" Then
        Word9 = Word8(RF)
    Else
        Word9 = Word2(LS) & " Crore " & Word7(RS)
    End If
End Function
Public Function Word10(intNUM10 As String) As String
    Dim LF As String
    Dim RF As String
    
    LF = Left$(intNUM10, 3)
    RF = Right$(intNUM10, 7)
    
    Word10 = Word3(LF) & " crore " & Word7(RF)
    
End Function

Public Function DecimalX(strValue As String) As Double
    Dim intCOUNT As Integer
    Dim i As Integer
    Dim strDECI As String
    
    intCOUNT = Len(Trim(strValue))
    For i = 1 To intCOUNT
        strDECI = Mid$(Trim(strValue), i, 1)
        If strDECI = "." Then
           DecimalX = i
           Exit Function
        End If
    Next i
End Function
Public Function SpaceX(strValue As String) As Integer
    Dim intCOUNT As Integer
    Dim i As Integer
    Dim intSPACE As Integer
    Dim strSPACE As String
    
    intCOUNT = Len(Trim(strValue))
    For i = 1 To intCOUNT
        strSPACE = Mid$(Trim(strValue), i, 1)
        If strSPACE = Space(1) Then
           SpaceX = i - 1
           Exit Function
        Else
           SpaceX = SpaceX + 1
        End If
    Next i
End Function
Public Function RightSpcX(strValue As String) As Integer
    Dim intCOUNT As Integer
    Dim i As Integer
    Dim intSPACE As Integer
    Dim strSPACE As String
    
    intCOUNT = Len(Trim(strValue))
    For i = intCOUNT To 1 Step -1
        strSPACE = Mid$(Trim(strValue), i, 1)
        If strSPACE = Space(1) Then
           RightSpcX = intCOUNT - i
           Exit Function
        Else
           RightSpcX = RightSpcX + 1
        End If
    Next i
End Function
