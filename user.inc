<%
If Session("User") = "" then
	 Session("User") = Request.cookies("UserName")
	 if Session("User") = "" then
	  Response.Redirect "TimeOut.asp"
		Response.end
	 end if
	 Session("CurrSSN") = Request.cookies("CurrSSN")
	 Session("IsClerk") = Request.cookies("IsClerk")
end if
%>
