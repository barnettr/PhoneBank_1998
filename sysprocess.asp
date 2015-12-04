<%
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DataConn")  'Use you own connection string here!

'Create a recordset and query sysprocesses
Dim rs
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT * FROM master..sysprocesses WHERE hostname <> ''", conn

'We query where hostname is NOT empty, so that we get the listing of
'user processes, not system processes.
%>

<HTML>
<BODY>
<TABLE border=1>
<TR>
<% For i = 0 to rs.Fields.Count - 1 %>
	<TH>
		<FONT SIZE=2><%=rs.Fields(i).Name%></FONT>
	</TH>
<%	Next %>

</TR>
<% Do While Not rs.EOF %>
<TR>
	<% For i = 0 to rs.Fields.Count - 1 %>
	<TD>
		<FONT SIZE=2><%=rs(i)%></FONT>
	</TD>
	<% Next %>	
</TR>
<%	rs.MoveNext
Loop %>
</TABLE>
</BODY>
</HTML>

<%
'Always important, clean up time!!
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
%>


