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
	<BR>
	<CENTER><font size="2" face="verdana, arial, helvetica">Do you wish to log this Phone Call?</font></CENTER>
	<BR>
	<HR>
	<BR>
	<TABLE WIDTH=100% BORDER=0>
		<TR>
			<TD ALIGN=CENTER>
				<BUTTON STYLE='width=100;' ID=Ok onClick='SaveLog(0)'>Yes</BUTTON>
				&nbsp;&nbsp;&nbsp;
				<BUTTON STYLE='width=100;' ID=Override onClick='SaveLog(1)'>No</BUTTON>
				&nbsp;&nbsp;&nbsp;
				<BUTTON STYLE='width=100;' ID=Cancel onClick='SaveLog(2)'>Cancel</BUTTON>
			</TD>
		</TR>
	</TABLE>
</BODY>
<SCRIPT LANGUAGE=vbscript>

	sub SaveLog(iIndex)
		select case iIndex
			case 0
				window.returnvalue = "normal"
			case 1
				window.returnvalue = "override"
			case 2
				window.returnvalue = "cancel"
		end select
		window.close
	end sub
</SCRIPT>
</HTML>
