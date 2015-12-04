<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<HTML>
<HEAD>


<script language="JavaScript" src='function.js'></script>

</HEAD>
<BODY>
<TABLE WIDTH=95% CELLPADDING=0 CELLSPACING=0 border="0">
	<TR>
		<TD ALIGN=CENTER>
			<A HREF='http://seaimg01/NWAHome' TARGET='_blank'><IMG SRC='images/home.gif' BORDER=0 ALT='NWA Home Page'></A>
		</TD>
	</TR>
	<TR>
		<TD align="center">
      <br>
			<font face="verdana, arial, helvetica" size="1"><b>Click the button to open your<br>Home Page</b></font>
		</TD>
	</TR>
</TABLE>

<!--- Name: <%=Session("NWAName")%>&nbsp;&nbsp;&nbsp;<INPUT TYPE=BUTTON VALUE=Cookie onClick='ShowCookie()'>
<BR>
UserID: <%=Session("User")%>&nbsp;&nbsp;&nbsp;<INPUT TYPE=BUTTON VALUE=CurrSSN onClick='ShowSSN()'>
<BR>
<A HREF=Abandon.asp>Abandon Session</A>
<BR>
<%
'If Session("Criteria") <> "" or Session("Criteria2") <> "" then
'	Response.Write "Current Criteria: <BR>"
'	If Session("Criteria") <> "" then 
'		Response.Write "&nbsp; " & Session("Criteria") & "<BR>"
'	end if
'	If Session("Criteria2") <> "" then 
'		Response.Write "&nbsp; " & Session("Criteria2") & "<BR>"
'	end if
'end if
'Response.Write "CurrSSN is " & top.NavFrame.iCurrSSN
%>
<BR> --->

</BODY>
<SCRIPT LANGUAGE=vbscript>

	sub ShowCookie
		msgbox "Cookie is <%= Request.cookies%>"
	end sub
	
	sub ShowSSN
		msgbox "CurrSSN is " & top.NavFrame.iCurrSSN
	end sub
	
</SCRIPT>
</HTML>
