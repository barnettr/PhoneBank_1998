<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<STYLE>TD.clsTab {
	BACKGROUND-COLOR: #003366; BORDER-BOTTOM: #99ccff 2px inset; BORDER-LEFT: #6699cc 1px solid; BORDER-RIGHT: #6699cc 1px solid; BORDER-TOP: #003366 2px solid; CURSOR: hand
}
TD.clsTab A {
	COLOR: #ffffff; FONT-SIZE: 95%; TEXT-DECORATION: none
}
TD.clsTab A:hover {
	COLOR: #ffffff; FONT-SIZE: 95%; TEXT-DECORATION: none
}
TD.clsTab A:active {
	COLOR: #ffffff; FONT-SIZE: 95%; TEXT-DECORATION: none
}
TD.clsTabSelected {
	BACKGROUND-COLOR: #6699cc; BORDER-LEFT: #99ccff 2px outset; BORDER-RIGHT: #99ccff 2px; BORDER-TOP: #99ccff 2px outset
}
TD.clsTabSelected A {
	COLOR: #ccffcc; FONT-SIZE: 95%; FONT-WEIGHT: bold; TEXT-DECORATION: none
}
TD.clsTabSelected A:hover {
	COLOR: #ccffcc; FONT-SIZE: 95%; FONT-WEIGHT: bold; TEXT-DECORATION: none
}
TD.clsTabSelected A:active {
	COLOR: #ccffcc; FONT-SIZE: 95%; FONT-WEIGHT: bold; TEXT-DECORATION: none
}
TABLE#idTabs TD {
	FONT-SIZE: 72%
}
TABLE#newsContent {
	MARGIN-RIGHT: 50px
}
</STYLE>
<SCRIPT language=JavaScript><!--
function TabClick( nTab )
{	nTab = parseInt(nTab);
	var oTab;
	var prevTab = nTab-1;
	var nextTab = nTab+1;
	event.cancelBubble = true;
	el = event.srcElement;
	for (var i = 0; i < newsContent.length; i++)
	{		oTab = tabs[i];
		oTab.className = "clsTab";
		oTab.style.borderLeftStyle = "";
		oTab.style.borderRightStyle = "";
		newsContent[i].style.display = "none";
	}
	newsContent[nTab].style.display = "block";
	tabs[nTab].className = "clsTabSelected";
	oTab = tabs[nextTab];
	if (oTab) oTab.style.borderLeftStyle = "none";
		oTab = tabs[prevTab];
	if (oTab) oTab.style.borderRightStyle = "none";
		event.returnValue = false;
	}
//--></SCRIPT>
<TABLE border=0 cellPadding=0 cellSpacing=0 height="100%" width="100%">
  <TR>
    <TD align=left vAlign=top width="95%">
      <TABLE bgColor=#6699cc border=0 cellPadding=2 cellSpacing=0 id=idTabs style="COLOR: #ffffff; DISPLAY: none">
        <TBODY>
        <TR height=25 vAlign=center>
          <TD class=clsTab id=tabs onclick="TabClick('0');">&nbsp;<A href="javscript:void();" onclick="return TabClick('0');"><font face="verdana, arial, helvetica">Participants</font></A>&nbsp;</TD>
          <TD class=clsTab id=tabs onclick="TabClick('1');">&nbsp;<A href="javscript:void();" onclick="return TabClick('1');"><font face="verdana, arial, helvetica">Providers</font></A>&nbsp;</TD>
          <TD class=clsTab id=tabs onclick="TabClick('2');">&nbsp;<A href="javscript:void();" onclick="return TabClick('2');"><font face="verdana, arial, helvetica">Checks</font></A>&nbsp;</TD>
          <TD class=clsTab id=tabs onclick="TabClick('3');">&nbsp;<A href="javscript:void();" onclick="return TabClick('3');"><font face="verdana, arial, helvetica">Other Searches</font></A>&nbsp;</TD>
        </TR><TR>
          <TD bgColor=#6699cc colSpan=6 height=16></TD>
        </TR>
        </TBODY>
      </TABLE>      

<SCRIPT>tabs[3].width="100%"; idTabs.style.display="block";</SCRIPT>

<TABLE border=0 cellPadding=0 cellSpacing=0 id="TabHead" width="100%">
  <TBODY>
  <TR>
    <TD align=left colSpan=2 height=25 width="100%"></TD>
  </TR>
  <TR>
    <TD bgColor=#336699 height=18 noWrap vAlign=center>&nbsp;&nbsp;<FONT color=#ffffff size=2 face="verdana, arial, helvetica"><B>Participants</B></FONT>&nbsp;&nbsp;</TD>
    <TD align=right width="100%"></TD>
  </TR>
  <TR>
    <TD align=left bgColor=#336699 colSpan=2 height=6 vAlign=top width="100%">&nbsp;</TD>
  </TR>
  </TBODY>
</TABLE>

<SCRIPT>TabHead.style.display = "none";</SCRIPT>

<TABLE border=0 cellPadding=1 cellSpacing=0 id="newsContent" width="92%" style="font:9pt verdana, arial, helvetica, sans-serif">
  <TBODY>
  <TR>
    <TD width=5></TD>
    <TD width="100%">The Big Columns</TD>
  </TR>
  </TBODY>
</TABLE>

<SCRIPT>newsContent.style.display="none";</SCRIPT>

<TABLE border=0 cellPadding=0 cellSpacing=0 id="TabHead" width="100%">
  <TBODY>
  <TR>
    <TD align=left colSpan=2 height=25 width="100%"></TD>
  </TR>
  <TR>
    <TD bgColor=#336699 height=18 noWrap vAlign=center>&nbsp;&nbsp;<FONT color=#ffffff size=2 face="verdana, arial, helvetica"><B>Providers</B></FONT>&nbsp;&nbsp;</TD>
    <TD align=right width="100%"></TD>
  </TR>
  <TR>
    <TD align=left bgColor=#336699 colSpan=2 height=6 vAlign=top width="100%">&nbsp;</TD>
  </TR>
  </TBODY>
</TABLE>

<SCRIPT>TabHead[1].style.display = "none";</SCRIPT>

<TABLE border=0 cellPadding=1 cellSpacing=0 id="newsContent" width="92%" style="font:9pt verdana, arial, helvetica, sans-serif">
  <TBODY>
  <TR>
    <TD width=5></TD><TD width="100%">The Columns</TD>
  </TR>
  </TBODY>
</TABLE>

<SCRIPT>newsContent[1].style.display="none";</SCRIPT>

<TABLE border=0 cellPadding=0 cellSpacing=0 id="TabHead" width="100%">
  <TBODY>
  <TR>
    <TD align=left colSpan=2 height=25 width="100%"></TD>
  </TR>
  <TR>
    <TD bgColor=#336699 height=18 noWrap vAlign=center>&nbsp;&nbsp;<FONT color=#ffffff size=2 face="verdana, arial, helvetica"><B>Checks</B></FONT>&nbsp;&nbsp;</TD>
    <TD align=right width="100%"></TD>
  </TR>
  <TR>
    <TD align=left bgColor=#336699 colSpan=2 height=6 vAlign=top width="100%">&nbsp;</TD>
  </TR>
  </TBODY>
</TABLE>

<SCRIPT>TabHead[2].style.display = "none";</SCRIPT>

<TABLE border=0 cellPadding=1 cellSpacing=0 id="newsContent" width="92%" style="font:9pt verdana, arial, helvetica, sans-serif">
  <TBODY>
  <TR>
    <TD width=5></TD>
    <TD width="100%">The Columns</TD>
  </TR>
  </TBODY>
</TABLE>

<SCRIPT>newsContent[2].style.display="none";</SCRIPT>

<TABLE border=0 cellPadding=0 cellSpacing=0 id="TabHead" width="100%">
  <TBODY>
  <TR>
    <TD align=left colSpan=2 height=25 width="100%"></TD>
  </TR>
  <TR>
    <TD bgColor=#336699 height=18 noWrap vAlign=center>&nbsp;&nbsp;<FONT color=#ffffff size=2 face="verdana, arial, helvetica"><B>Other Searches</B></FONT>&nbsp;&nbsp;</TD>
    <TD align=right width="100%"></TD>
  </TR>
  <TR>
    <TD align=left bgColor=#336699 colSpan=2 height=6 vAlign=top width="100%">&nbsp;</TD>
  </TR>
  </TBODY>
</TABLE>

<SCRIPT>TabHead[3].style.display = "none";</SCRIPT>

<TABLE border=0 cellPadding=1 cellSpacing=0 id="newsContent" width="92%" style="font:9pt verdana, arial, helvetica, sans-serif">
  <TBODY>
  <TR>
    <TD width=5></TD>
    <TD width="100%">The Columns</TD>
  </TR>
  </TBODY>
</TABLE>

<SCRIPT>newsContent[3].style.display="none";</SCRIPT>
</TABLE>
<SCRIPT>newsContent[0].style.display = "block";tabs[0].className = "clsTabSelected";</SCRIPT>


</body>
</html>
