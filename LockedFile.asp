<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%  
Dim SSN, Name, Reason
Dim adoConn, adoRS, sSQL, LastName, DepNo, locktype

SSN = request.querystring("SSN")
Name = request.querystring("Name")
LastName = request.querystring("LastName")
DepNo = request.querystring("DepNo")
locktype = request.querystring("locktype")

'response.write locktype

set adoConn = Server.CreateObject("ADODB.Connection")
adoConn.CommandTimeout = 0
set adoRS = Server.CreateObject("ADODB.Recordset")
adoConn.Open Application("DataConn")
if LastName = "" then
  sSQL = "select LockedReason, LockedBy, LockedDate from Names where SSN=" & SSN 
end if
if SSN = "" then
  sSQL = "select LockedReason LockedBy, LockedDate from Names where LastName='" & LastName & "'"
end if 
adoRS.Open sSQL, adoConn, adOpenKeyset, adLockOptimistic
%>


  <html>
  <head>
  <!--#include file="VBFuncs.inc" -->
  </head>
  
  <body onload="UpdateScreen(3)">
  <font size="2" face="verdana, arial, helvetica">
  <p>&nbsp;
  <p>&nbsp;
  <center><p><font size="4" face="arial, helvetica" color="red"><b>Access Denied !!</b></font>
  <p><b>Administrative Lock On PDF. Contact Your Supervisor.</b>
  <p><b>To return to a new search please click the Participants/Dependent Tab.</b>
  <p><b>This file is locked because:
  <p><font color="#983300"><% if adoRS("LockedReason") <> "" then %><%= adoRS("LockedReason") %><% else %>NO REASON IN DATABASE!!<% end if %></font>
  <p>Locked By: <font color="#983300"><% if adoRS("LockedBy") <> "" then %><%= adoRS("LockedBy") %><% else %>NO NAME IN DATABASE!!<% end if %></font>
  <p>Locked Date: <font color="#983300"><% if adoRS("LockedDate") <> "" then %><%= formatdatetime(adoRS("LockedDate"),vbShortDate) %><% else %>NO DATE IN DATABASE!!<% end if %></font>
  <p>The Lock Types are <font color="#983300"><%= locktype %></font>
  <p>This participants SSN is <font color="#983300"><%= SSN %></font>. Please log this call.</b>
  <p><img SRC="images/log.gif" onClick="LogCall()">
  
  
  </center>
  </font>

<%  
adoRS.Close
adoConn.Close
set adoRS = nothing
set adoConn = nothing
%> 
  </body>
  
  </html>
