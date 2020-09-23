Attribute VB_Name = "CreatexMyDB"
Option Explicit

' ========================================================
' === Generator       : CreateMDB v1.00.0012
' === Copyright©      : 2000-2001 NiKroWare
' === Created         : 7/1/2002 1:50:59 PM
' === Access Database : xMyDB.mdb
' ========================================================

Private CAT As ADOX.Catalog

Public Sub CreateMDB(ByVal dbPathFilename As String)
  On Error GoTo CreateERROR
  Dim CAT As ADOX.Catalog
  Dim TBL As ADOX.Table
  Dim INDX As ADOX.Index
  Dim CMD As ADODB.Command

  '/* Engine Type = 4; (Access97)
  '/* Engine Type = 5; (Access2000)

    Set CAT = New ADOX.Catalog
    CAT.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & dbPathFilename & ";" & _
               "Jet OLEDB:Database Password=;" & _
               "Jet OLEDB:Engine Type=4;"
    
    
    '/* Create Table 'MyTable' */
    Set TBL = New ADOX.Table
    Set TBL.ParentCatalog = CAT
    With TBL
        .Name = "MyTable"
        .Columns.Append "UID", adInteger, 0
        .Columns("UID").Properties("AutoIncrement") = True
        .Columns("UID").Properties("NullAble") = True
        .Columns.Append "ByteField", adUnsignedTinyInt, 0
        .Columns("ByteField").Properties("NullAble") = True
        .Columns.Append "IntegerField", adSmallInt, 0
        .Columns("IntegerField").Properties("NullAble") = True
        .Columns.Append "LongField", adInteger, 0
        .Columns("LongField").Properties("NullAble") = True
        .Columns.Append "SingleField", adSingle, 0
        .Columns("SingleField").Properties("NullAble") = True
        .Columns.Append "DoubleField", adDouble, 0
        .Columns("DoubleField").Properties("NullAble") = True
        .Columns.Append "CurrencyField", adCurrency, 0
        .Columns("CurrencyField").Properties("NullAble") = True
        .Columns.Append "TextField", adVarWChar, 50
        .Columns("TextField").Properties("NullAble") = True
        .Columns("TextField").Properties("Jet OLEDB:Allow Zero Length") = True
        .Columns.Append "BooleanField", adBoolean, 2
        .Columns.Append "DateTimeField", adDate, 0
        .Columns("DateTimeField").Properties("NullAble") = True
        .Columns.Append "MemoField", adLongVarWChar, 0
        .Columns("MemoField").Properties("NullAble") = True
        .Columns("MemoField").Properties("Jet OLEDB:Allow Zero Length") = True
    End With
    CAT.Tables.Append TBL
    
    
    '/* Create Index 'PrimaryKey' */
    Set INDX = New ADOX.Index
    With INDX
        .Name = "PrimaryKey"
        .Columns.Append "UID"
        .PrimaryKey = True
        .Unique = True
        .Clustered = False
        .IndexNulls = adIndexNullsDisallow
    End With
    CAT.Tables("MyTable").Indexes.Append INDX
    Set INDX = Nothing
    

    '/* Create Query 'MyQuery' */
    Set CMD = New ADODB.Command
    CMD.CommandText = "SELECT MyTable.UID, MyTable.ByteField, MyTable.IntegerField, MyTable.LongField, " & _
          "MyTable.SingleField, MyTable.DoubleField, MyTable.CurrencyField, MyTable.TextField, " & _
          "MyTable.BooleanField, MyTable.DateTimeField, MyTable.MemoField FROM MyTable " & _
          "ORDER BY MyTable.ByteField;"
    CAT.Views.Append "MyQuery", CMD
    Set CMD = Nothing

    
    Set CAT = Nothing

Exit Sub

CreateERROR:

End Sub

