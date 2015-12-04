<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<HTML>
<HEAD>
<TITLE></TITLE>
</HEAD>
<BODY>
	<BR>
	<CENTER><font face="verdana, arial, helvetica" size="2" color="red"><b>Enter an SSN to log a call for, then click the Ok button</b></font></CENTER>
	<BR>
	<HR>
	<BR>
	&nbsp;&nbsp;&nbsp;SSN: <INPUT ID=SSN TYPE=TEXT SIZE=10 VALUE=''>
	<BR>
	<BR>
	<BR>
	<TABLE WIDTH=100% BORDER=0>
		<TR>
			<TD ALIGN=CENTER>
				<BUTTON STYLE='width=100;' ID=Ok onClick='SaveLog(0)'>Ok</BUTTON>
				&nbsp;&nbsp;&nbsp;
				<BUTTON STYLE='width=100;' ID=Cancel onClick='SaveLog(1)'>Cancel</BUTTON>
			</TD>
		</TR>
	</TABLE>
</BODY>
<SCRIPT LANGUAGE=vbscript>

	sub SaveLog(iIndex)
		select case iIndex
			case 0
				if not isNumeric(SSN.value) then
					msgbox "Please enter a valid SSN.",,"Invalid Data"
					exit sub
				end if
				window.returnvalue = trim(SSN.value)
			case 1
				window.returnvalue = "cancel"
		end select
		window.close
	end sub
</SCRIPT>
</HTML>
