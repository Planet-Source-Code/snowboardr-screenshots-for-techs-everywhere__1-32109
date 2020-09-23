<%@ Language=VBScript %>
<%


LinkAndFolder = "http://68.6.129.137:81/ss/"  'server address and any extra directories you want to specify must have forwardslash at the end "/"
HomePage = "home.asp?page=home"    	      'Home Page to display to user


Background = "on" 'If off it will hide the background color menu from user


image = "on"      'If off the user can not view pictures by Image, they can only view ss by text only & text / URL
textURL = "on"	  'If off the user can not view text / URL


Response.write "<html>"			
							
Response.write "<!--%" & LinkAndFolder & "%-->"		
Response.write "<!--background:" & Background & "--!>"  
Response.write "<!--image:" & image & "--!>"       
Response.write "<!--textURL:" & textURL & "--!>"     


Response.write "</html>"



'Response.redirect LinkAndFolder & HomePage


%>






