<div align="center">

## Display SQL Results


</div>

### Description

Display a page that allows you to type in a SQL Query statement, and display the results in HTML Table format.

The number of field parameters and field names are immaterial. It will display the results in an easy to read format. Great for customer reports.

The Code will not allow an Update or Delete query.
 
### More Info
 
You will have to provide the connection to your database. There are examples in the code.

An HTML Page


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Eugene](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/eugene.md)
**Level**          |Advanced
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/eugene-display-sql-results__4-6474/archive/master.zip)





### Source Code

```
<%
' Open Connection to the database
dim Connect
set Connect = Server.CreateObject("ADODB.Connection")
REM - Un-Rem one of the connect statements depending on your data
'	connection method, and change the phyisical path if necessary
'1. ODBC
'strConn = "DSN=ASP-Forum; uid=;password=;"
'2. SQL-OLEDB
'strConn = "PROVIDER=SQLOLEDB; uid=;password=;Initial Catalog=ASPForum;"
'3. ADO-Access
'strConn = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=d:\inetpub\wwwroot\asp\forum\forum.mdb"
'4. OLE 4.0
'strConn = "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=d:\inetpub\wwwroot\asp\forum\forum.mdb"
Connect.Open strConn
sSQL = request.form("Sequel")
%>
<HTML>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<table width=600 border=0 cellpadding=0 cellspacing=0>
<tr align="LEFT" valign="TOP">
<!--
***********************************************************************
Make sure change the Action statement if you change the name of the ASP.
-->
<form action="Viewlist.asp" method=POST>
<td>
	SQL Statement:<br>
	<TEXTAREA cols=60 name=Sequel rows=8 wrap=PHYSICAL><%=sSQL%></TEXTAREA>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<INPUT name=SubmitButton type=submit value=Submit>
	&nbsp;&nbsp;
	<INPUT type=reset value=Reset><BR>
</td>
</form>
</tr>
</table>
<%
if left(ucase(trim(sSQL)), 6) = "UPDATE" or left(ucase(trim(sSQL)), 6) = "DELETE" or left(ucase(trim(sSQL)), 4) = "DROP" then
	%>
	<table width=600 border=0 cellpadding=0 cellspacing=0>
		<tr>
			<td>
				<br>
				<font face=Verdana,Arial size=+0><b>Function Not Allowed from this view</b></font>
				<br>
			</td>
		</tr>
	</table>
	<%
elseif sSQL > "" then
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sSQL, Connect
	flds = rs.Fields.Count
	if not (rs.bof and rs.eof) then
		%>
		<table width=600 border=0 cellpadding=0 cellspacing=0>
			<tr>
				<td>
					<br>
					<font face=Verdana,Arial size=+0><b>Searching for selected data...</b></font>
					<br>
				</td>
			</tr>
		</table>
		<%
		dim sColor(1)
		toggle = 1
		sColor(0) = ""
		sColor(1) = "bgcolor=#eeeeee"
		response.write "<TABLE width=600 border=0 cellpadding=3 cellspacing=0>" & vbCrLf
		response.write "<TR bgcolor=#FAF0E6>" & vbCrLf
		For i = 0 To flds - 1
			response.write "<TD>" & rs.Fields(i).Name & "</TD>"
		Next
		response.write vbCrLf & "</TR>"
		rs.movefirst
		Do While Not rs.EOF
			toggle = 1 - toggle
			response.write "<TR " & sColor(toggle) & ">" & vbCrLf
			For i = 0 To flds - 1
				response.write "<TD>" & rs.Fields(i).Value & "&nbsp;</TD>"
			Next
			response.write vbCrLf & "</TR>"
			rs.MoveNext
		Loop
	end if
	rs.Close
	Set rs = Nothing
	response.write vbcrlf & "</TABLE>"
end if
%>
</html>
```

