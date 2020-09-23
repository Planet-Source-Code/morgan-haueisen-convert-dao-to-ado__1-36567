VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsADOConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  Dim OpenADO As New clsADOConnect
'  Dim MyDB As New ADODB.Connection
'  Dim MySet As New ADODB.Recordset
'
'/* to open a database and recordset use
'
'    OpenADO.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, DBname
'    MySet.open "SELECT Location.* FROM Location " & OrderBy, MyDB, adOpenStatic,adLockPessimistic
'
'/* or
'
' Call OpenDB(MyDB, ,DBname)
' Call OpenRS(MySet, "SELECT Location.* FROM Location", MyDB)

Public Enum e_DatabaseTypes
    dbt_Undefined = 0
    dbt_OracleMSDA = 1
    dbt_OracleODBC = 2
    dbt_SQLserver = 3
    dbt_MicrosoftAccess97File = 4
    dbt_MicrosoftAccess2KFile = 5
    dbt_MicrosoftAccessFile = 6
    dbt_DSNFile = 7
End Enum

Public Function CompressAccessDatabase(mSourceDB As String, Optional Password As String = "", Optional DB_Type As e_DatabaseTypes = dbt_MicrosoftAccess2KFile) As Boolean
  Dim srcDB As String
  Dim destDB As String
  Dim JRO As JRO.JetEngine
  Dim EngineType As String

    On Error GoTo Err
    
    If DB_Type = dbt_MicrosoftAccess2KFile Then
        EngineType = "5"
    Else
        EngineType = "4"
    End If
    
    srcDB = mSourceDB
    destDB = App.Path & "\MHbackup.mdb"
    If Dir(destDB) > vbNullString Then Kill destDB
    
    Set JRO = New JRO.JetEngine
    JRO.CompactDatabase _
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & srcDB & ";Jet OLEDB:Database Password=" & Password & ";", _
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & destDB & ";Jet OLEDB:Database Password=" & Password & ";Jet OLEDB:Engine Type=" & EngineType & ";"

    Kill srcDB

    DoEvents
    Name destDB As srcDB
    
    Set JRO = Nothing
    CompressAccessDatabase = True
Exit Function

Err:
    CompressAccessDatabase = False
    Set JRO = Nothing
End Function

Public Function BuildConnectString(ByVal DatabaseType As e_DatabaseTypes, ByVal ServerOrFilename As String, Optional ByVal DatabaseName As String, Optional ByVal UserName As String, Optional ByVal Password As String) As String
    Select Case DatabaseType
        Case dbt_OracleMSDA
            BuildConnectString = "Provider=MSDAORA;Data Source=" & ServerOrFilename & _
                                 ";User ID=" & IIf(UserName <> "", UserName, "") & _
                                 ";Password=" & IIf(Password <> "", Password, "") & ";" & _
                                 IIf(DatabaseName <> "", "Initial Catalog=" & DatabaseName & ";", "")
        Case dbt_OracleODBC
            BuildConnectString = "DRIVER={Microsoft ODBC for Oracle};SERVER=" & ServerOrFilename & _
                                 ";UID=" & UserName & ";PWD=" & Password & ";" & _
                                 IIf(DatabaseName <> "", "Initial Catalog=" & DatabaseName & ";", "")
        Case dbt_SQLserver
            BuildConnectString = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & ServerOrFilename & _
                                  ";User ID=" & IIf(UserName <> "", UserName, "") & _
                                  ";Password=" & IIf(Password <> "", Password, "") & ";" & _
                                  IIf(DatabaseName <> "", "Initial Catalog=" & DatabaseName & ";", "")
        Case dbt_DSNFile
            BuildConnectString = "Provider=MSDASQL;DSN=" & ServerOrFilename & ";UID=" & _
                                 IIf(UserName <> "", UserName, "") & _
                                 ";PWD=" & IIf(Password <> "", Password & ";", "") & ";" & _
                                 IIf(DatabaseName <> "", "Initial Catalog=" & DatabaseName & ";", "")
                                 
        Case dbt_MicrosoftAccess2KFile
            BuildConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                 "Data Source=" & ServerOrFilename & ";" & _
                                 "Jet OLEDB:Database Password=" & Password & ";" & _
                                 "Jet OLEDB:Engine Type=5;"
                                 
        Case dbt_MicrosoftAccess97File
            BuildConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                 "Data Source=" & ServerOrFilename & ";" & _
                                 "Jet OLEDB:Database Password=" & Password & ";" & _
                                 "Jet OLEDB:Engine Type=4;"
                                 
        Case dbt_MicrosoftAccessFile
            BuildConnectString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & ServerOrFilename & _
                                 ";DefaultDir=" & DetermineDirectory(ServerOrFilename) & _
                                 ";PWD=" & IIf(Password <> "", Password & ";", ";")
    End Select
    
End Function

Private Function DetermineDirectory(inputString As String) As String
  Dim i As Integer
    i = InStrRev(inputString, "\", , vbTextCompare)
    DetermineDirectory = Mid(inputString, 1, i)
End Function

Public Function adoConnectOpen(oADODBConnection As ADODB.Connection, _
                                DB_Type As e_DatabaseTypes, _
                                ByVal ServerOrFilename As String, _
                                Optional ByVal dbPathFilename As String = "", _
                                Optional CommandType As CommandTypeEnum = adCmdStoredProc, _
                                Optional CursorLocation As CursorLocationEnum = adUseClient, _
                                Optional ByVal UserName As String = "", _
                                Optional ByVal Password As String = "") As Boolean
    
    
    On Error GoTo adoConnectError
    
    Set oADODBConnection = New ADODB.Connection
    
    If dbPathFilename = vbNullString Then dbPathFilename = ServerOrFilename

    With oADODBConnection
        .CursorLocation = CursorLocation '/* default=adUseClient(3)
        .Open BuildConnectString(DB_Type, ServerOrFilename, dbPathFilename, UserName, Password)
    End With

    adoConnectOpen = True
    
Exit Function
    
    
adoConnectError:
    MsgBox "Connect To database failed with: " & Err.Number & " " & Err.Description
    adoConnectOpen = False
    
End Function
Public Function adoConnectString(DB_Type As e_DatabaseTypes, _
                                ByVal ServerOrFilename As String, _
                                Optional ByVal dbPathFilename As String = "", _
                                Optional ByVal UserName As String = "", _
                                Optional ByVal Password As String = "") As String
    
    If dbPathFilename = vbNullString Then dbPathFilename = ServerOrFilename
    adoConnectString = BuildConnectString(DB_Type, ServerOrFilename, dbPathFilename, UserName, Password)
    
End Function


Public Function adoConnectClose(oADODBConnection As ADODB.Connection) As Boolean
    
    On Error GoTo adoConnectError
    oADODBConnection.Close
    adoConnectClose = True
    
Exit Function
    
    
adoConnectError:
    MsgBox "Connect To database failed with: " & Err.Number & " " & Err.Description
    adoConnectClose = False
    
End Function



