<%@ LANGUAGE=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim sPath,i
Dim adoConn, adoRS, adoCmd, adoParam, sSQL, sTemp
'response.write "Session User is = " Session("User")

	set adoConn = Server.CreateObject("ADODB.Connection")
'**********************************************************************************
' Setting CommandTimeout to 0 waits indefinitely for response; note that per 
' ADO documentation Command object will not inherit the Connection setting (default
' is 30 seconds).  This appeared to help with response issues on seasql02.
'**********************************************************************************
	adoConn.CommandTimeout = 0
	set adoCmd = Server.CreateObject("ADODB.Command")
	set adoRS = Server.CreateObject("ADODB.Recordset")
 
	adoConn.Open Application("AdminConn")
	adoCmd.CommandText = "web_userinfo"
	adoCmd.CommandType = adCmdStoredProc
	Set adoCmd.ActiveConnection = adoConn
	'*************************************************************************************************************************************
	' For CreateParameter method, parameters are:
	'	[Name As String], [Type As DataTypeEnum = adEmpty], [Direction As ParameterDirectionEnum = adParamInput], [Size As Long], [Value])
	' and the value for the SP are:
	'	user, adVarChar, adParamInput, 50 (field is varChar 50), username (Session("User"))
	'*************************************************************************************************************************************
	Set adoParam = adoCmd.CreateParameter("@user", 200, 1, 50, Session("User"))
	adoCmd.Parameters.Append adoParam
	adoRS.Open adoCmd
	    
	if adoRS.EOF then
		set adoConn = nothing
		set adoCmd = nothing
		set adoRS = nothing
		Response.Redirect "notindb.asp"
	else
		sPath=""
		sTemp = Request.ServerVariables("PATH_INFO")
		i=instr(sTemp,"/")
		do while i>0
			sPath=sPath & left(sTemp,i)
			sTemp=mid(sTemp,i+1,len(sTemp))
			i=instr(sTemp,"/")
		loop
		Session("HTTPRootPath")=sPath
		sPath=""
		sTemp = Request("Path_Translated")
		i=instr(sTemp,"\")
		do while i>0
			sPath=sPath & left(sTemp,i)
			sTemp=mid(sTemp,i+1,len(sTemp))
			i=instr(sTemp,"\")
		loop
		Session("RootPath")=sPath
		Session("NWAName") = adoRS("firstname") & " " & adoRS("lastname")
		Session("JobCatID") = trim(adoRS("jobcatid"))
	end if
	adoRS.Close
	set adoCmd = nothing
	adoConn.Close
'	set adoConn = nothing
'	set adoRS = nothing

'	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
'**********************************************************************************
' The following will be used to determine if a user should have the IsClerk
' flag set -- it requires the field to be added to the UserInformation
' table, and then the value to be set appropriately for each user.  Until then
' the value is being set to true manually.
'**********************************************************************************
	sSQL = "select * from UserInformation where LogonID='" & Session("User") & "'"
	adoRS.Open sSQL, adoConn, adOpenKeyset, adLockOptimistic
'**********************************************************************************
' The following is temporary, until the IsClerk flag is
' implemented in UserInformation, to allow both Clerk and
' non-Clerk operation to be observed.
'**********************************************************************************
	if Request.QueryString("NotClerk") <> "" then
		Session("IsClerk") = false
		Response.Cookies("IsClerk") = Session("IsClerk")
	else
		Session("IsClerk") = true
		Response.Cookies("IsClerk") = Session("IsClerk")
	end if
	
	set adoConn = nothing
	set adoRS = nothing
%>

<HTML>
<HEAD>
<TITLE>Phone Bank (Claimstest Database)</TITLE>		
</HEAD>
	<FRAMESET FRAMESPACING=0 ROWS="15%,*">
		<FRAMESET FRAMESPACING=0 COLS="80%,*">
			<FRAME SRC="NavMain.asp" NAME="NavFrame" FRAMEBORDER=NO SCROLLING=NO>
			<FRAME SRC="Status.asp" NAME="AppStatus" SCROLLING=NO NORESIZE FRAMEBORDER=NO>
		</FRAMESET>
		<FRAME SRC="Welcome.htm" NAME="Details" FRAMEBORDER=NO>
	</FRAMESET>
</HTML>


