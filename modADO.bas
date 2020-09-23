Attribute VB_Name = "modADO"
'/*************************************/
'/* Author: Morgan Haueisen
'/* Copyright (c) 1997-2002
'/*************************************/

Option Explicit

'/* For password protected database file (if required) */
Public Const DB_PWD As String = "morgan"
Public Const DB_Type As String = "5" '/* 4=Access97; 5=Access2000

Public Type goUserType
    UserName As String
    Password As String
    MachineName As String
End Type
Public goUser As goUserType

Public Type goAppType
    SourceDB As String
    SecurityFile As String
    SystemID As String
End Type
Public goApp As goAppType

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function xGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Function GetUserName() As String
 Dim strBuffer As String * 255
 Dim lngBufferLength As Long
 Dim lngRet As Long
 Dim strTemp As String, i As Integer

    lngBufferLength = 255
    lngRet = xGetUserName(strBuffer, lngBufferLength)
    strTemp = LCase$(strBuffer)
    GetUserName = ClipNull(strTemp)
    
End Function
Private Function ClipNull(InString As String) As String
  Dim intpos As Integer
    If Len(InString) Then
        intpos = InStr(InString, vbNullChar)
        If intpos > 0 Then
            ClipNull = left(InString, intpos - 1)
        Else
            ClipNull = InString
        End If
    End If
End Function

Public Function ADOFindFirst(MySet As ADODB.Recordset, ByVal Filter As String) As Boolean
  Dim mhRS As ADODB.Recordset
  Dim mhMatch As Boolean

    Set mhRS = New ADODB.Recordset
    Set mhRS = MySet.Clone
    mhRS.Filter = Filter
    
    If mhRS.RecordCount > 0 Then
        mhRS.MoveFirst
        MySet.Bookmark = mhRS.Bookmark
        mhMatch = True
    Else
        If MySet.RecordCount > 0 Then
            MySet.MoveLast
            MySet.MoveNext
        End If
        mhMatch = False
    End If
    
    mhRS.Close
    Set mhRS = Nothing
    
    ADOFindFirst = mhMatch

End Function

Public Function ADOFindNext(MySet As ADODB.Recordset, ByVal Filter As String) As Boolean
  Dim mhRS As ADODB.Recordset
  Dim mhMatch As Boolean

    Set mhRS = New ADODB.Recordset
    Set mhRS = MySet.Clone
    mhRS.Filter = Filter
    mhRS.Sort = MySet.Sort
    
    If mhRS.RecordCount > 0 Then
        mhRS.Bookmark = MySet.Bookmark
        mhRS.MoveNext
        If Not mhRS.EOF Then
            MySet.Bookmark = mhRS.Bookmark
            mhMatch = True
        Else
            mhMatch = False
        End If
    Else
        If MySet.RecordCount > 0 Then
            MySet.MoveLast: MySet.MoveNext
        End If
        mhMatch = False
    End If
    
    mhRS.Close
    Set mhRS = Nothing
    ADOFindNext = mhMatch
    
End Function

Public Sub ADOCreateQuery(ByVal sFilename As String, ByVal QueryName As String, ByVal QueryString As String, Optional AddDelay As Boolean = False)
  Dim OpenConnect As New clsADOConnect
  Dim CAT As New ADOX.Catalog
  Dim CMD As New ADODB.Command

    On Local Error Resume Next
    '/* Open the catalog
    CAT.ActiveConnection = OpenConnect.adoConnectString(dbt_MicrosoftAccess2KFile, sFilename, sFilename, , DB_PWD)
    
    '/* Create the query
    CMD.CommandText = QueryString
    CAT.Views.Append QueryName, CMD
    DoEvents
    
    Set CAT = Nothing
    'Set CMD = Nothing
    'Set OpenConnect = Nothing
    On Local Error GoTo 0
    DoEvents
    If AddDelay Then Sleep 5000

End Sub

Public Sub ADODeleteQuery(ByVal sFilename As String, ByVal QueryName As String)
  Dim OpenConnect As New clsADOConnect
  Dim CAT As New ADOX.Catalog
  Dim CMD As New ADODB.Command

    '/* Open the catalog
    CAT.ActiveConnection = OpenConnect.adoConnectString(dbt_MicrosoftAccess2KFile, sFilename, sFilename, , DB_PWD)
    '/* Delete the query
    CAT.Views.Delete QueryName
    
    Set CAT = Nothing
    Set CMD = Nothing
    Set OpenConnect = Nothing

End Sub


Public Sub ADOAttachTable(TableName As String, ByVal AttachFromMDB As String, ByVal AttachToMDB As String)
  Dim OpenConnect As New clsADOConnect
  Dim CAT As New ADOX.Catalog
  Dim TBL As New ADOX.Table

   On Local Error Resume Next
   '/* Open the catalog
   CAT.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AttachFromMDB & ";"

   '/* Set the name and target catalog for the table
   TBL.Name = TableName
   Set TBL.ParentCatalog = CAT

   '/* Set the properties to create the link
   TBL.Properties("Jet OLEDB:Create Link") = True
   TBL.Properties("Jet OLEDB:Link Datasource") = AttachToMDB
   TBL.Properties("Jet OLEDB:Link Provider String") = ";Pwd=" & DB_PWD
   TBL.Properties("Jet OLEDB:Remote Table Name") = TableName

   '/* Append the table to the collection
   CAT.Tables.Append TBL

   Set CAT = Nothing
   Set TBL = Nothing
   Set OpenConnect = Nothing

End Sub

Public Sub InitSettings(ByVal MDBfile As String, Optional SECfile As String = "x", Optional SYSID As String = "")
    
    goApp.SourceDB = MDBfile
    goApp.SecurityFile = SECfile
    goApp.SystemID = "MLH" & SYSID
    
    goUser.MachineName = Environ("computername")
    goUser.UserName = GetUserName
    
    If Dir$(MDBfile) = vbNullString Then
        MsgBox "The Database file is missing.  Please contact your system adminstrator for assistance", vbCritical
        End
    End If
    
    If SECfile <> "x" Then
        If Dir$(SECfile) = vbNullString Then
            MsgBox "The Security file is missing.  Please contact your system adminstrator for assistance", vbCritical
            End
        End If
    End If
    
End Sub

Public Sub OpenDB(MyDB As ADODB.Connection, Optional ByVal OpenMDB As Boolean = True, Optional ByVal DBPathName As String = vbNullString)
  Dim OpenConnect As clsADOConnect
  
    If DBPathName = vbNullString Then
        DBPathName = goApp.SourceDB
    End If
    
    If OpenMDB Then
        '/* Password protected database file */
        Set OpenConnect = New clsADOConnect
        OpenConnect.adoConnectOpen MyDB, dbt_MicrosoftAccess2KFile, DBPathName, DBPathName, , , , DB_PWD
        Set OpenConnect = Nothing
    Else
        MyDB.Close
        Set MyDB = Nothing
    End If
    DoEvents
End Sub

Public Sub OpenRS(oActiveRecordset As ADODB.Recordset, ByVal oSourceTable As String, oActiveConnection As ADODB.Connection, Optional oCursorType As CursorTypeEnum = adOpenStatic, Optional oLockType As LockTypeEnum = adLockOptimistic, Optional ByVal oOptions As Integer = -1)
    Set oActiveRecordset = New ADODB.Recordset
    oActiveRecordset.Open oSourceTable, oActiveConnection, oCursorType, oLockType, oOptions
End Sub
Public Function ADOFindPrevious(MySet As ADODB.Recordset, ByVal Filter As String) As Boolean
  Dim mhRS As ADODB.Recordset
  Dim mhMatch As Boolean

    Set mhRS = New ADODB.Recordset
    Set mhRS = MySet.Clone
    mhRS.Filter = Filter
    mhRS.Sort = MySet.Sort
    
    If mhRS.RecordCount > 0 Then
        mhRS.Bookmark = MySet.Bookmark
        mhRS.MovePrevious
        If (Not mhRS.BOF) Then
            MySet.Bookmark = mhRS.Bookmark
            mhMatch = True
        Else
            mhMatch = False
        End If
    Else
        If MySet.RecordCount > 0 Then
            MySet.MoveFirst: MySet.MovePrevious
        End If
        mhMatch = False
    End If
    
    mhRS.Close
    Set mhRS = Nothing

    ADOFindPrevious = mhMatch
    
End Function

Public Function ADOFindLast(MySet As ADODB.Recordset, ByVal Filter As String) As Boolean
  Dim mhRS As ADODB.Recordset
  Dim mhMatch As Boolean

    Set mhRS = New ADODB.Recordset
    Set mhRS = MySet.Clone
    mhRS.Filter = Filter
    
    If mhRS.RecordCount > 0 Then
        mhRS.MoveLast
        MySet.Bookmark = mhRS.Bookmark
        mhMatch = True
    Else
        If MySet.RecordCount > 0 Then
            MySet.MoveLast
            MySet.MoveNext
        End If
        mhMatch = False
    End If
    
    mhRS.Close
    Set mhRS = Nothing
    ADOFindLast = mhMatch

End Function

Public Sub ADODeleteTable(sFilename As String, sTableName As String)
  On Error GoTo ErrTrapD
    Dim OpenConnect As New clsADOConnect
    Dim CAT As ADOX.Catalog
    Set CAT = New ADOX.Catalog
    
    '/* Open Database
    CAT.ActiveConnection = OpenConnect.adoConnectString(dbt_MicrosoftAccess2KFile, sFilename, sFilename, , DB_PWD)
        
    '"Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & sFilename & ";" & _
               "Jet OLEDB:Database Password=" & DB_PWD & ";" & _
               "Jet OLEDB:Engine Type=5;"
    
    '/* Delete table
    Dim TBL As ADOX.Table
    Set TBL = New ADOX.Table
    TBL.Name = sTableName
    Set TBL.ParentCatalog = CAT
    CAT.Tables.Delete TBL
    
    Set OpenConnect = Nothing
    Set TBL = Nothing
    Set CAT = Nothing
Exit Sub

ErrTrapD:
  'MsgBox Err.Number & " / " & Err.Description
  Exit Sub
  Resume

End Sub

