<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<HTML>
<HEAD>
<TITLE>Northwest Administrators</TITLE>
</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 RIGHTMARGIN=0 onLoad='SwitchCritTabs(-1)'>
<BR>
<TABLE WIDTH=100% CELLPADDING=0 CELLSPACING=0 BORDER=0>
	<TR>
		<TD ALIGN=CENTER VALIGN=CENTER>
			<font face="verdana, arial, helvetica" size="2"><b>Select one of the Search types below, and enter your search criteria.</b></font>
		</TD>
	</TR>
</TABLE>

<TABLE width=760 HEIGHT=26 STYLE='font-size: 8pt; cursor:hand;' CELLPADDING=0 CELLSPACING=0 BORDER=0>
	<TR>
		<TD WIDTH=190 ALIGN=CENTER VALIGN=CENTER style='font-size: 8pt;' background='images/entab3.gif' id=search name=search myNorm=images/entab3.gif myDis=images/distab2.gif onClick='SwitchCritTabs(0)'><font color="white" face="verdana, arial, helvetica"><b>
			Participants/Dependents</b></font>
		</TD>
		<TD WIDTH=190 ALIGN=CENTER VALIGN=CENTER background='images/distab2.gif' id=search name=search myNorm=images/entab3.gif myDis=images/distab2.gif onClick='SwitchCritTabs(1)'><font color="white" face="verdana, arial, helvetica"><b>
			Providers</b></font></TD>
		<TD WIDTH=190 ALIGN=CENTER VALIGN=CENTER background='images/distab2.gif' id=search name=search myNorm=images/entab3.gif myDis=images/distab2.gif onClick='SwitchCritTabs(2)'><font color="white" face="verdana, arial, helvetica"><b>
			Checks</b></font></TD>
		<TD WIDTH=190 ALIGN=CENTER VALIGN=CENTER background='images/distab2.gif' id=search name=search myNorm=images/entab3.gif myDis=images/distab2.gif onClick='SwitchCritTabs(3)'><font color="white" face="verdana, arial, helvetica"><b>
			Other Searches</b></font>
		</TD>
	</TR>
</TABLE>
<DIV ID=CritDiv STYLE='display:none;'>
	<TABLE BGCOLOR=#63659C WIDTH=100% CELLPADDING=2 CELLSPACING=2>
		<TR>
			<TD><font color="white" face="verdana, arial, helvetica" size="2">
				Use this section to search for details for Participants/Dependents:
			</font></TD>
		</TR>
	</TABLE>
</DIV>
<DIV ID=CritDiv STYLE='display:none;'>
	<TABLE BGCOLOR=#63659C WIDTH=100% CELLPADDING=2 CELLSPACING=2>
		<TR>
			<TD><font color="white" face="verdana, arial, helvetica" size="2">
				Use this section to search for details for Providers:
			</font></TD>
		</TR>
	</TABLE>
</DIV>
<DIV ID=CritDiv STYLE='display:none;'>
	<TABLE BGCOLOR=#63659C WIDTH=100% CELLPADDING=2 CELLSPACING=2>
		<TR>
			<TD><font color="white" face="verdana, arial, helvetica" size="2">
				Use this section to search for details for Checks:
			</font></TD>
		</TR>
	</TABLE>
</DIV>
<DIV ID=CritDiv STYLE='display:none;'>
	<TABLE BGCOLOR=#63659C WIDTH=100% CELLPADDING=2 CELLSPACING=2>
		<TR>
			<TD ALIGN=CENTER WIDTH=25%>
				<INPUT TYPE=BUTTON ID=Search VALUE='Phone Calls' STYLE='cursor:hand;' onClick='PickButton(3)'>
			</TD>
			<!--- <TD ALIGN=CENTER WIDTH=25%>
				<INPUT TYPE=BUTTON ID=Search VALUE='Claims' STYLE='cursor:hand;' onClick='PickButton(4)'>
			</TD> --->
			<TD ALIGN=CENTER WIDTH=25%>
				<INPUT TYPE=BUTTON ID=Search VALUE='ACE Letters' STYLE='cursor:hand;' onClick='PickButton(5)'>
			</TD>
			<TD ALIGN=CENTER WIDTH=25%>
				<INPUT TYPE=BUTTON ID=Search VALUE='ACS Letters' STYLE='cursor:hand;' onClick='PickButton(6)'>
			</TD>
		</TR>
	</TABLE>
</DIV>
<FORM ID=SearchPost TARGET=Details ACTION='' METHOD=GET>
	<INPUT TYPE=HIDDEN ID=Crit NAME=UserCriteria VALUE=''>
	<INPUT TYPE=HIDDEN ID=PDCriteria NAME=PDCriteria VALUE=''>
	<INPUT TYPE=HIDDEN ID=DepSSN NAME=UseDepSSN VALUE=''>
	<INPUT TYPE=HIDDEN ID=TaxID NAME=TaxID VALUE=''>
	<INPUT TYPE=HIDDEN ID=AltTaxID NAME=UseAltTaxID VALUE=''>
	<INPUT TYPE=HIDDEN ID=CheckAcct NAME=CheckAcct VALUE=''>
	<INPUT TYPE=HIDDEN ID=CheckNum NAME=CheckNum VALUE=''>
</FORM>
<!--#include file="VBFuncs.inc" -->
<SCRIPT LANGUAGE="VBScript">

	dim iCurrSSN
	iCurrSSN = -1	
	dim iCurrDiv
	iCurrDiv=0
	
	sub PickButton(iIndex)
		select case iIndex
			case 3 
				WorkingStatus
				SearchPost.Action = "PhoneSearch.asp"
				SearchPost.Crit.value = "Initial"
				SearchPost.submit
			case 4 
				WorkingStatus
				SearchPost.Action = "ClaimSearch.asp"
				SearchPost.Crit.value = "Initial"
				SearchPost.submit
			case 5 
				WorkingStatus
				SearchPost.Action = "FormLetterSearch.asp"
				SearchPost.Crit.value = "Initial"
				SearchPost.submit
			case 6 
				WorkingStatus
				SearchPost.Action = "LetterSearch.asp"
				SearchPost.Crit.value = "Initial"
				SearchPost.submit
			case else
				msgbox "none"
		end select
	end sub
	
	sub SwitchCritTabs(iIndex)

		if iIndex = -1 then
			document.all.item("CritDiv",iCurrDiv).style.display="block"
			document.all.item("Search",iCurrDiv).background=document.all.item("Search",iCurrDiv).myNorm
			document.all.item("Search",iCurrDiv).style.fontsize="8pt"
			WorkingStatus
			SearchPost.Action = "PersonSearch.asp"
			SearchPost.Crit.value = "Initial"
			SearchPost.submit
		else		
			ChangeCritTabVisual(iIndex)
			if iIndex = 3 then
				top.Details.location.href = "Other.asp"
			else
				WorkingStatus
				select case iIndex
					case 0
						SearchPost.Action = "PersonSearch.asp"
					case 1 
						SearchPost.Action = "ProviderSearch.asp"
					case 2 
						SearchPost.Action = "CheckSearch.asp"
				end select
				SearchPost.Crit.value = "Initial"
				SearchPost.submit
			end if
		end if
	end sub
	
	sub ChangeCritTabVisual(iIndex)
	
		if iIndex = iCurrDiv then
			exit sub
		end if
	
		if iCurrDiv <> -1 then
			document.all.item("CritDiv",iCurrDiv).style.display = "none"
			document.all.item("Search",iCurrDiv).background = document.all.item("Search",iCurrDiv).myDis
			document.all.item("Search",iCurrDiv).style.fontsize = "8pt"
      document.all.item("Search",iCurrDiv).style.textdecoration = "none"
		end if
		iCurrDiv = iIndex
		document.all.item("CritDiv",iCurrDiv).style.display = "block"
		document.all.item("Search",iCurrDiv).background = document.all.item("Search",iCurrDiv).myNorm
		document.all.item("Search",iCurrDiv).style.fontsize = "8pt"
    document.all.item("Search",iCurrDiv).style.textdecoration = "underline"
	
	end sub

</SCRIPT>
</BODY>
</HTML>



