
<%
ServerAddy = "http://localhost/ss/"

Dim img
img = Request.QueryString("img")
FileDes = Request.QueryString("Information")

Response.Write "<br>" & FileDes & "<br>"
Response.Write "<img src=" & Chr(34) & ServerAddy & img &  Chr(34) & ">"
Response.end

%>
