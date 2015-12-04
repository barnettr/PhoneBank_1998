<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
dim adoConn, adoRS, adoCmd, adoParam, sSQL, sTemp
dim i, m_bContinue
dim m_sEASFund, m_sEASPlan, m_sEASRate
dim m_sACSFund, m_sACSPlan, m_sACSCat, m_sACSDesc

	m_bContinue = true
	
	m_sEASFund = Request.QueryString("EASFund")
	m_sEASPlan = Request.QueryString("EASPlan")
	m_sEASRate = Request.QueryString("EASRate")
	
	set adoConn = Server.CreateObject("ADODB.Connection")
	adoConn.CommandTimeout = 0
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
   
  sSQL = "PB_GetACSFundPlan " &"'"& m_sEASFund & "'" & ", " & "'" & m_sEASPlan & "'" & ", " & "'" & m_sEASRate & "'"
	
	Set adoRS = adoConn.execute(sSQL)
  if not adoRS.EOF then
		m_sACSFund = adoRS("ACSFund")
		m_sACSPlan = adoRS("ACSPlan")
  else
    adoRS.Close
  end if

  sSQL = "FindPlan " & "'" & m_sACSFund & "'" & ", " & "'" & m_sACSPlan & "'"		
		
  Set adoRS = adoConn.execute(sSQL)
  if not adoRS.EOF then
	  m_sACSCat = adoRS("PlanCat")
		m_sACSDesc = adoRS("Description")
  else
    adoRS.Close
		adoConn.Close
		set adoRS = nothing
		set adoConn = nothing
  end if
	
%>
<HTML>
<HEAD>
<TITLE>Details</TITLE>
</HEAD>
<BODY TOPMARGIN=2 LEFTMARGIN=2 RIGHTMARGIN=0 LANGUAGE=VBScript>
<!--<INPUT type=button onClick='javascript:ChangeTitle()' id=button1 name=button1>-->
	<FONT SIZE=2 face="verdana, arial, helvetica"><CENTER><b>ACS Plan Description</b></CENTER></FONT>
	<HR>
  <p>&nbsp;	
	<TABLE WIDTH=100% COLS=5 CELLPADDING=1 CELLSPACING=1 style="color:black; font:9pt verdana, arial, helvetica, sans-serif">
		<TR>
			<TD><b>EAS Fund:</b></TD>
			<TD><b><%= m_sEASFund%></b></TD>
			<TD WIDTH=5%></TD>
			<TD><b>ACS Fund:</b></TD>
			<TD><b><% if m_sACSFund <> "" then %><%= m_sACSFund%><% else %><font color="red">No ACS Fund Available</font><% end if %></b></TD>
		</TR>
		<TR>
			<TD><b>EAS Plan:</b></TD>
      <TD><b><%= m_sEASPlan%></b></TD>
			<TD WIDTH=5%></TD>
			<TD><b>ACS Plan Code:</b></TD>
			<TD><b><% if m_sACSPlan <> "" then %><%= m_sACSPlan%><% else %><font color="red">No ACS Plan Available</font><% end if %></b></TD>
		</TR>
		<TR>
			<TD><b>EAS Rate:</b></TD>
			<TD><b><%= m_sEASRate%></b></TD>
			<TD WIDTH=5%></TD>
			<TD><b>ACS Plan Cat.:</b></TD>
			<TD><b><% if m_sACSCat <> "" then %><%= m_sACSCat%><% else %><font color="red">No ACS Plan Cat Available</font><% end if %></b></TD>
		</TR>
	</TABLE>
  <p>&nbsp;
	<TABLE WIDTH=100% COLS=2 CELLPADDING=1 CELLSPACING=1 style="color:black; font:9pt verdana, arial, helvetica, sans-serif">
		<TR>
			<TD VALIGN=TOP><b>Description:</b></TD>
			<TD><b><% if m_sACSDesc <> "" then %><%= m_sACSDesc%><% else %><font color="red">No Description Available</font><% end if %></b></TD>
		</TR>
		<TR>
			<TD ALIGN=CENTER COLSPAN=6>&nbsp;</TD>
		</TR>
		<TR>
			<TD ALIGN=CENTER COLSPAN=2><INPUT TYPE=BUTTON VALUE='OK' onClick='javascript:Done()'></TD>
		</TR>
	</TABLE>

</BODY>
<SCRIPT LANGUAGE=vbscript>

	sub Done()
		window.close
	end sub

	sub ChangeTitle()
	msgbox "in ChangeTitle"
	top.document.title = "My Title"
	end sub
</SCRIPT>
</HTML>
