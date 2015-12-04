<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<% 
Dim Claim
Claim = request.querystring("ClaimNum")
%>

<html>
<head><title>AceLink Test (ExecuteWinApp)</title>
<!--#include file="VBFuncs.inc" -->
</head>
<body bgcolor=white text=black onLoad="AceDouble()">

<object ID = "AceLink"
 classID = "CLSID:31C03DD3-7BF7-11D3-ACA5-00600837EAC7"
 codebase = "AceLink/setup.exe#version=1,0,0,0">
</object>

<script language="vbscript">
  sub Worksheets()
    Dim Result
    Dim oObject
    Set oObject = CreateObject("AceLink.cAceLink")
    oObject.executable = "C:\ACE\Worksheets.exe"
    oObject.Server = "databases.nwadmin.com"
    oObject.Database = "Claimstest"
    oObject.ClaimNumber = "<%= Claim %>"
    Result = oObject.RunApp
  end sub
  
  sub AceDouble()
    UpdateScreen(3)
    Worksheets()
  end sub
</script>

  <blockquote>
  <center>
    <h3><font face="arial">Worksheets Application Interface</font></h3>
    <p><img SRC="images/bluebar2.gif" onClick="history.go(-1)">
    <p><font face="verdana, arial, helvetica" size="2">The back button will take you back to the Claim Search page. If you need to re-visit the Worksheets for this particular Claim Number please click the "Go To Worksheets" button below.</font>
    <p><input type="button" name="worksheets" id="worksheets" value="Go To Worksheets" onClick="Worksheets()">
  </center>
  </blockquote>

</body>
</html>
