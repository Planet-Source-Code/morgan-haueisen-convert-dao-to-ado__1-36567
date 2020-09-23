<div align="center">

## Convert DAO to ADO


</div>

### Description

This will give you the tools and instructions necessary to convert your DAO project to ADO.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-07-05 08:04:56
**By**             |[Morgan Haueisen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/morgan-haueisen.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Convert\_DA102221752002\.zip](https://github.com/Planet-Source-Code/morgan-haueisen-convert-dao-to-ado__1-36567/archive/master.zip)





### Source Code

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<META NAME="Generator" CONTENT="Microsoft Word 97">
<TITLE>To convert your DAO project just follow these steps</TITLE>
</HEAD>
<BODY LINK="#0000ff">
<FONT SIZE=2>
<B><P>Here are 9 steps to convert your DAO project to ADO.</B> I have included version numbers but yours may vary.</P>
<P> I have also included a sample BAS file that will create an Access database file. Use this file as a model to create your own or you can download some nice code that will make this task much easier.</P>
<P>See </FONT><A HREF="http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=22085&lngWId=1)"><FONT SIZE=2>http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=22085&lngWId=1</FONT></A></P>
<FONT SIZE=2>
<B><P>1. Add the following references to your project:</P>
<UL>
</B><LI>Microsoft ActiveX Data Objects 2.6 Library</LI>
<LI>Microsoft ADO Ext. 2.5 for DDL and Security</LI>
<LI>Microsoft Jet and Replication Objects 2.5 Library</LI></UL>
<B><P>2. Remove references to Microsoft DAO 3.6 Object Library</P>
</B>
<B><P>3. Add the component Microsoft ADO Data Control (if required)</P>
</B>
<B><P>4. Replace any data bound controls with their ADO equivalents</P><DIR>
</B><P>Examples:</P></DIR>
<UL>
<LI>Microsoft Data Bound Grid with Microsoft DataGrid Control (OLEDB)</LI>
<LI>Microsoft Data Bound List with Microsoft DataList Control (OLEDB)</LI></UL>
<OL START=5>
<B><LI>Add the attached BAS and CLS files to your application.</LI></OL>
<UL>
</B><LI>clsADOConnect.cls</LI>
<LI>modADO.bas</LI>
<LI>modADOdc.bas (only needed if you are using the ADO Data control)</LI></UL>
<B><P>6. Search your project for the following and replace as shown:</P>
<UL>
</B><LI>Dim MyDB As Database WITH Dim MyDB As ADODB.Connection</LI>
<LI>Dim MySet As Recordset WITH Dim MySet As ADODB.Recordset</LI>
<LI>Set MyDB = Workspaces(0).Opendatabase(dbFileName) WITH Call OpenDB(MyDB, , dbFileName)</LI>
<LI>Set MySet = MyDB.Openrecordset("Select * From Table") WITH Call OpenRS(MySet, "Select * From Table", MyDB)</LI></UL>
<UL>
<LI>Data1. DatabaseName= dbFileName AND Data1. RecordSource WITH ADOdcConnect(Data1, "Select * From Table", dbFileName)</LI></UL>
<DIR>
<DIR>
<P>The Following is true for FindFIrst, FindNext, FindLast, FindPrevious:</P></DIR>
</DIR>
<UL>
<LI>MySet.FindFirst "[LastName]=’Haueisen’" WITH ADOFindFirst(MySet, "[LastName]=’Haueisen’")</LI>
<LI>If MySet.NoMatch then WITH If Not ADOFindFirst(MySet, "[FieldName]=’Morgan’") then</LI></UL>
<DIR>
<DIR>
<P>The Following is true for FindFIrst, FindNext, FindLast, FindPrevious:</P></DIR>
</DIR>
<UL>
<LI>Data1.Recordset.FindFirst "[LastName]=’Haueisen’" WITH ADOdcFindFirst(MySet, "[LastName]=’Haueisen’")</LI>
<LI>If Data1.Recordset.NoMatch then WITH If Not ADOdcFindFirst(MySet, "[FieldName]=’Morgan’") then</LI></UL>
<UL>
<LI>Move your code from Data1_Reposition() to Data1_MoveComplete()</LI>
<LI>Move your code from Data1_Validate () to Data1_WillChangeRecord () and modify as necessary (Save = False to adStatus = adStatusCancel). </LI></UL>
<UL>
<LI>If you are using queries such as "Select * From Table Where [LastName] LIKE ‘Ha*’" you need to change them to look like "Select * From Table Where [LastName] LIKE ‘Ha%’"</LI></UL>
<B><P>7. Remove all MySet.Edit</P>
<OL START=8>
<LI>Use the following as needed:</LI></OL>
<UL>
</B><LI>ADOAttachTable</LI>
<LI>ADOCreateQuery</LI>
<LI>ADODeleteQuery</LI>
<LI>InitSettings (does some system stuff)</LI>
<LI>GetUserName (Returns a user’s login ID)</LI></UL>
<B>
<P>9. Test., Test, Test.</P>
<UL>
</B><LI>ADO follows the SQL rules for writing queries; so you may need to make some changes to any queries you have imbedded in your code. For example the word Size is a key word and this is the reason for the last line under step 6.</LI></UL>
<B></B></FONT></BODY>
</HTML>

