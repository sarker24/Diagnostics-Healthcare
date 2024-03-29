VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private str_Connect As String
Private str_Provider As String
Private str_Server As String
Private str_Database As String
Private str_User As String
Private str_Pass As String

Public Property Get Connection() As String
Dim snc As String
Dim FileNumber As Integer

FileNumber = FreeFile

Open App.Path + "\LOCATE.dat" For Input Access Read As #FileNumber
Input #FileNumber, str_Provider
Input #FileNumber, str_Server
Input #FileNumber, str_Database
Input #FileNumber, str_User
Input #FileNumber, str_Pass
Close #FileNumber

        If str_Provider <> "" Then
            snc = "Provider = " + str_Provider
        Else
            snc = "Provider=SQLOLEDB"
        End If

        If str_Pass <> "" Then
            snc = snc + ";Password=" + str_Pass + ";Trusted_Connection=Yes"
        Else
            snc = snc + ";Trusted_Connection=Yes"
        End If

        If str_User <> "" Then
            snc = snc + ";User ID=" + str_User
        Else
            snc = snc + ";User ID= Sa"
        End If
        
        If str_Database <> "" Then
            snc = snc + ";Initial Catalog=" + str_Database
        Else
            snc = snc + ";Initial Catalog= Pubs"
        End If
        
        If str_Server <> "" Then
            snc = snc + ";Data Source =" + str_Server
        Else
            snc = snc + ";Data Source = ALPS"
        End If
        Connection = snc + ";time out=0"
End Property

Public Property Let Connection(ByVal snc As String)
    str_Connect = snc
End Property

Public Property Get Provider() As String
    Provider = str_Provider
End Property

Public Property Let Provider(ByVal sNewProvider As String)
    str_Provider = sNewProvider
End Property

Public Property Get Database() As String
    Database = str_Database
End Property

Public Property Let Database(ByVal sNewDatabase As String)
    str_Database = sNewDatabase
End Property

Public Property Get User() As String
    User = str_User
End Property

Public Property Let User(ByVal sNewUser As String)
    str_User = sNewUser
End Property

Public Property Get Pass() As String
    Pass = str_Pass
End Property

Public Property Let Pass(ByVal sNewPass As String)
    str_Pass = sNewPass
End Property

Public Property Get Server() As String
    Server = str_Server
End Property

Public Property Let Server(ByVal sNewServer As String)
    str_Server = sNewServer
End Property

