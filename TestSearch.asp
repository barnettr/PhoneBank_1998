<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<% Response.Buffer = true %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim adoConn, adoRS, sSQL, sTemp, sCriteria
Dim sPageData, iPageNo, iRowCount, i, iPageCount
Dim m_iSSN, m_sLastName, m_sAddress1, m_sCity, m_sState, m_sZip
dim m_bUseDepSSN, m_bUseEAS
dim bUseSQL, bContinueProcessing

bContinueProcessing = true

	function PSearchFailed()
	dim sRedir

		sRedir = "PSearchFailed.asp?"
		sRedir = sRedir & "SSN=" & m_iSSN
		sRedir = sRedir & "&Address1=" & m_sAddress1
		sRedir = sRedir & "&City=" & m_sCity
		sRedir = sRedir & "&LastName=" & m_sLastName
		sRedir = sRedir & "&State=" & m_sState
		sRedir = sRedir & "&ZIP=" & m_sZip
		sRedir = sRedir & "&UseDepSSN=" & m_bUseDepSSN
		sRedir = sRedir & "&UseEAS=" & m_bUseEAS
		
		Response.Redirect(sRedir)
		Response.End
	end function

If Request.Querystring("UserCriteria") = "Initial" then
	bUseSQL = false
Else
	bUseSQL = true
End If
	Response.Write "<BR>"
	Response.Write "Request.ServerVariables(QUERY_STRING)" & Request.ServerVariables("QUERY_STRING")  & "<BR>"
	Response.Write "Request.QueryString(UserCriteria)=" & Request.QueryString("UserCriteria")  & "<BR>"
	Response.Write "Request.Querystring=" & Request.Querystring  & "<BR>"
	Response.Write "Request.Querystring(UserCriteria)=" & Request.Querystring("UserCriteria")  & "<BR>"
%>
<html>
<head>
<title>Details</title>
</head>
<body TOPMARGIN="2" LEFTMARGIN="2" RIGHTMARGIN="0" LANGUAGE="VBScript" onLoad="UpdateScreen(0)">
<link REL="STYLESHEET" HREF="styles/CritTable.css">

<%


%>
</body>
<!--#include file="VBFuncs.inc" -->
<script LANGUAGE="VBSCRIPT">

	sub ValidSend(iIndex)
	dim bValidData, iTemp

		Criteria.SSN.Value = trim(Criteria.SSN.Value)
		Criteria.LastName.Value = trim(Criteria.LastName.Value)
		Criteria.Address1.Value = trim(Criteria.Address1.Value)
		Criteria.City.Value = trim(Criteria.City.Value)
		Criteria.State.Value = trim(Criteria.State.Value)
		Criteria.Zip.Value = trim(Criteria.Zip.Value)
		bValidData = false
		if Criteria.SSN.Value <> ""  then
			if not isNumeric(Criteria.SSN.Value) then
				msgbox "Please enter a valid SSN.",,"Invalid Data"
				exit sub
			else
				if Criteria.SSN.Value > <%= Application("IntMax")%> or _
						Criteria.SSN.Value < <%= Application("IntMin")%> then
					msgbox "Please enter a valid SSN.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.IncludeDepSSN.Checked and Criteria.SSN.Value = "" then
			msgbox "To search by Dependent SSN, you have to enter a valid SSN.",,"Invalid Data"
			exit sub
		end if
		if Criteria.LastName.Value <> "" then
			if ContainsInvalids(Criteria.LastName.Value) then
				msgbox "Please remove invalid characters from the Last Name field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Address1.Value <> "" then
			if ContainsInvalids(Criteria.Address1.Value) then
				msgbox "Please remove invalid characters from the Address field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.City.Value <> "" then
			if ContainsInvalids(Criteria.City.Value) then
				msgbox "Please remove invalid characters from the City field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.State.Value <> "" then
			if ContainsInvalids(Criteria.State.Value) then
				msgbox "Please remove invalid characters from the State field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Zip.Value <> "" then
			iTemp = replace(Criteria.Zip.Value,"-","")
			if not isNumeric(iTemp) then
				msgbox "Please enter a valid Zip Code.",,"Invalid Data"
				exit sub
			else
				if iTemp > <%= Application("IntMax")%> or _
						iTemp < <%= Application("IntMin")%> then
					msgbox "Please enter a valid Zip Code.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if not bValidData then
			msgbox "You haven't entered anything to search for.",,"No Criteria"
			exit sub
		end if
		
		WorkingStatus
		if "<%= m_iSSN%>" <> Criteria.SSN.Value or _
				"<%= m_sLastName%>" <> Criteria.LastName.Value or _
				"<%= m_sAddress1%>" <> Criteria.Address1.Value or _
				"<%= m_sCity%>" <> Criteria.City.Value or _
				"<%= m_sState%>" <> Criteria.State.Value or _
				"<%= m_sZip%>" <> Criteria.Zip.Value then
			Criteria.ActionType.value=""
		else
			Criteria.ActionType.value=iIndex
		end if
		Criteria.submit
	end sub
		
</script>
</html>
