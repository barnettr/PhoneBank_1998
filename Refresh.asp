<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
<!--#include file="adovbs_mod.inc" -->
	<title>Untitled</title>
</head>

<body>
<%  
  set adoConn = Server.CreateObject("ADODB.Connection")
  adoConn.ConnectionTimeout = 0
  adoConn.CommandTimeout = 0
  set adoCmd = Server.CreateObject("ADODB.Command")
  adoCmd.CommandTimeout = 0
  set adoRS = Server.CreateObject("ADODB.Recordset")
  adoConn.Open Application("DataConn")
  
  adoCmd.CommandText = "PB_GetRxClaims"
	adoCmd.CommandType = &H0004
	Set adoCmd.ActiveConnection = adoConn
  adoCmd.Parameters.Refresh

%>
<table border=1>
<caption>Parameter Information</caption>
  <tr>
    <th>Parameter Name</th>
    <th>Datatype</th>
    <th>Direction</th>
    <th>Size</th>
  </tr>
  <% for each thing in adoCmd.Parameters %>
  <tr>
    <td><%= thing.name %></td>
    <td><%= thing.type %></td>
    <td><%= thing.direction %></td>
    <td><%= thing.size %></td>
  </tr>
<%   next
     adoConn.Close 
%>
</table>

</body>
</html>
