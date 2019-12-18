	<%
	Function RequestEncode(byval input)
		if Request.Querystring(input) <> "" then
			RequestEncode=Request.Querystring(input)
		elseif Request.Form(input) <> "" then
			RequestEncode=Request.Form(input)
		elseif Request.Servervariables(input) <> "" then
			RequestEncode=Request.Servervariables(input)
		else
			RequestEncode=Request(input)
		end if
		RequestEncode = server.HTMLencode(RequestEncode)
		RequestEncode = replace(RequestEncode,"&amp;","&")
	End Function

	Function SessionC (byval input)
		On Error Resume Next
		Set SessionC = Session(input)
		If Err.number<>0 then
			SessionC=Session(input)
			Err.clear
		End If
        if LCASE(input) = "sessionid" then SessionC = session.sessionid
	End Function
	
	%>