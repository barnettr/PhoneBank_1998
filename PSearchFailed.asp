<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<%
Dim m_iSSN, m_sLastName, m_sAddress1, m_sCity, m_sState, m_sZip
dim m_bUseDepSSN, m_bUseEAS, OtherArea

	m_bUseEAS = Request.QueryString("UseEAS")
	m_bUseDepSSN =  Request.QueryString("UseDepSSN")
	m_iSSN = Request.QueryString("SSN")
	m_sLastName = Request.QueryString("LastName")
	m_sAddress1 = Request.QueryString("Address1")
	m_sCity = Request.QueryString("City")
	m_sState = Request.QueryString("State")
	m_sZip = Request.QueryString("ZIP")
  OtherArea = request.querystring("OtherArea")
%>
<!--#include file="user.inc" -->
<html>
<head>
<title></title>
</head>
<body TOPMARGIN="2" LEFTMARGIN="2" RIGHTMARGIN="0" LANGUAGE="VBScript" onLoad="UpdateStatus()">
	<table WIDTH="100%" COLS="3">
		<tr>
			<td ALIGN="CENTER" WIDTH="10%">
			</td>
			<td ALIGN="CENTER">
				<font SIZE="+2" face="verdana, arial, helvetica"><strong>Participant/Dependent Search</strong></font>
			</td>
			<td ALIGN="CENTER" WIDTH="20%">
				<img SRC="images/bluebar2.gif" onClick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0">
			</td>
		</tr>
	</table>
	<br>
	<center><font face="verdana, arial, helvetica" size="2"><b>Current Criteria</b></font>
  <p>&nbsp;
	<table WIDTH="75%" ALIGN="CENTER" COLS="4" border="0" style="color:black; font:9pt verdana, arial, helvetica, sans-serif">
		<tr>
			<td WIDTH="25%" ALIGN="RIGHT"><b>SSN:</b></td>
			<td WIDTH="25%" STYLE='color:<%= Session("EmphColor")%>;'><%= m_iSSN%></td>
			<td WIDTH="25%" ALIGN="RIGHT"><b>Address:</b></td>
			<td WIDTH="25%" STYLE='color:<%= Session("EmphColor")%>;'><%= m_sAddress1%></td>
		</tr>
		<tr>
			<td ALIGN="RIGHT"><b>Last Name:</b></td>
			<td STYLE='color:<%= Session("EmphColor")%>;'><%= m_sLastName%></td>
			<td ALIGN="RIGHT"><b>City:</b></td>
			<td STYLE='color:<%= Session("EmphColor")%>;'><%= m_sCity%></td>
		</tr>
		<tr>
			<td ALIGN="RIGHT"><b>Search for Dep. SSN:</b></td>
			<td STYLE='color:<%= Session("EmphColor")%>;'><%= m_bUseDepSSN%></td>
			<td ALIGN="RIGHT"><b>State:</b></td>
			<td STYLE='color:<%= Session("EmphColor")%>;'><%= m_sState%></td>
		</tr>
		<tr>
			<td ALIGN="RIGHT"><b>Use EAS Data:</b></td>
			<td STYLE='color:<%= Session("EmphColor")%>;'><%= m_bUseEAS%></td>
			<td ALIGN="RIGHT"><b>ZIP:</b></td>
			<td STYLE='color:<%= Session("EmphColor")%>;'><%= m_sZip%></td>
		</tr>
      <td ALIGN="RIGHT"><b>Other Area Data:</b></td>
      <td STYLE='color:<%= Session("EmphColor")%>;'><%= OtherArea%></td>
    <tr>
    </tr>
	</table>
	</center>
	<br>
	<center><font color="red" face="verdana, arial, helvetica" size="2"><b>No matches were found.</b></font></center>
	<hr>
	<table WIDTH="100%" BORDER="0">
<%
		if m_iSSN <> "" then
%>
			<tr>
				<td>
					<button STYLE="width=200;" ID="PSearch" onClick="SetResult(0)">Log Call</button>
				</td>
				<td>
					<font face="verdana, arial, helvetica" size="2"><b>Log this phone call anyway.</b></font>
				</td>
			</tr>
<%
		end if
		if not m_bUseDepSSN and m_iSSN <> "" then
%>
			<tr>
				<td>
					<button STYLE="width=200;" ID="PSearch" onClick="SetResult(1)">Redo w/Dep. SSN</button>
				</td>
				<td><font color="red" face="verdana, arial, helvetica" size="2">
					<b>Re-try the search, also checking for Dependent SSN, (slow search; EAS data not used).</b>
				</font></td>
			</tr>
<%
		end if
		if not m_bUseEAS then
%>
			<tr>
				<td>
					<button STYLE="width=200;" ID="PSearch" onClick="SetResult(2)">Redo w/EAS</button>
				</td>
				<td><font color="red" face="verdana, arial, helvetica" size="2">
					<b>Re-try the search using EAS data (but don't search for Dep. SSN).</b>
				</font></td>
			</tr>
<%
		end if
		if (not m_bUseEAS or not m_bUseDepSSN) and m_iSSN <> "" then
%>
			<tr>
				<td>
					<button STYLE="width=200;" ID="PSearch" onClick="SetResult(3)">Redo w/Dep. SSN &amp; EAS</button>
				</td>
				<td><font color="red" face="verdana, arial, helvetica" size="2">
					<b>Re-try the search checking for Dependent SSN &amp; using EAS data (slow search).</b>
				</font></td>
			</tr>
<%
		end if
    if m_bUseEAS then
      if not OtherArea then
%>
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td colspan="2"><font color="red" face="verdana, arial, helvetica" size="2"><b>Please try all the other search options first. This is the last search you should try from this page.</b></font></td>
    </tr>
		<tr>
			<td>
        <button STYLE="width=200;" ID="PSearch" onClick="SetResult(4)">Other Area Search</button>
			</td>
			<td><font color="red" face="verdana, arial, helvetica" size="2">
				<b>Other Area Search.</b>
			</font></td>
		</tr>
<%  
   end if
 end if  
%>    
	</table>
	
	<form ID="Criteria" METHOD="GET" ACTION="PersonSearch.asp" TARGET="Details">
		<input TYPE="HIDDEN" ID="SSN" NAME="SSN" SIZE="10" VALUE="<%= m_iSSN%>">
		<input TYPE="HIDDEN" ID="Address1" NAME="Address1" SIZE="20" VALUE="<%= m_sAddress1%>">
		<input TYPE="HIDDEN" ID="City" NAME="City" SIZE="15" VALUE="<%= m_sCity%>">
		<input TYPE="HIDDEN" ID="LastName" NAME="LastName" SIZE="20" VALUE="<%= m_sLastName%>">
		<input TYPE="HIDDEN" ID="State" NAME="State" SIZE="3" VALUE="<%= m_sState%>">
		<input TYPE="HIDDEN" ID="ZIP" NAME="ZIP" SIZE="12" VALUE="<%= m_sZip%>">
		<input TYPE="HIDDEN" ID="DepSSN" NAME="UseDepSSN" VALUE>
		<input TYPE="HIDDEN" ID="UseEAS" NAME="UseEAS" VALUE>
    <input TYPE="HIDDEN" ID="OtherArea" NAME="OtherArea" VALUE>
	</form>
</body>
<!--#include file="VBFuncs.inc" -->
<script LANGUAGE="vbscript">

	sub SetResult(iIndex)
	
		WorkingStatus
		select case iIndex
			case 0
				location.replace("WriteLog.asp?SSN=<%= m_iSSN%>&OpenPhoneSearch=true&LogType=normal")
			case 1
				Criteria.UseDepSSN.value = true
				Criteria.UseEAS.value = false
				Criteria.submit
			case 2
				Criteria.UseDepSSN.value = false
				Criteria.UseEAS.value = true
				Criteria.submit
			case 3
				Criteria.UseDepSSN.value = true
				Criteria.UseEAS.value = true
				Criteria.submit
			case 4
        'Criteria.UseDepSSN.value = false
				Criteria.OtherArea.value = true
				Criteria.submit
'				history.back
		end select
	end sub
</script>
</html>
