<html>
<head>
<title>Screen Shots</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#CCCCCC" text="#000000">
<p><font face="Arial, Helvetica, sans-serif"><b><font size="5" color="#666666">S</font><font size="5" color="#999999"> 
  c r e e n <font color="#666666">S</font> h o t <font size="1">beta test</font></font></b></font></p>

<% 

If Request.Querystring("page") = "404" then
	Response.Write "<font face='Arial, Helvetica, sans-serif'><b><font size='3' color='#000000'>" & "This page cannont be found!" & "</font>"

ElseIf Request.Querystring("page") = "500" then
	Response.Write "<font face='Arial, Helvetica, sans-serif'><b><font size='3' color='#000000'>" &"There was an error getting your request" & "</font>"
	

Elseif Request.Querystring("page") = "home" then
	Response.Write "<font face='Arial, Helvetica, sans-serif'><b><font size='2' color='#000000'>" & "Home page here" & "</font>"

	End if



%>


</body>
</html>
