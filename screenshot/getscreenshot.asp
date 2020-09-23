<%@ Language=VBScript %>

<% 
	
	ServerDirectory = "F:\inetpub\wwwroot\"    'root directory
  	ServerAddy = "http://68.6.129.137:81/"  'server address 
	ExtraDir = "ss"  		           'extra Directories
    	

	Dim objFileScripting, objFolder
    	Dim filename, filecollection, strDirectoryPath, strUrlPath
   	Dim NoDash, FileDes, Folder, URL, Area
 	Dim ArrowPic, fcolor, mfcolor
	Dim PassWordProtect, NoSlash, MainFolder
    	
	
	Folder = Request.QueryString("Folder")	'Folder in link
    	URL = Replace(Folder, "\" , "/")	'Change folder slashes to URL slashes \ /
	Area = Request.QueryString("Area")   'Where the site is at
	color = Request.QueryString("color")  'Background color
	
	


	strDirectoryPath= ServerDirectory & ExtraDir & " \" & Folder   'Remove [ & "\" ] if there is no Extra Directory
    	

	strUrlPath = ServerAddy & ExtraDir & "/" & URL   
    

	Set objFileScripting = CreateObject("Scripting.FileSystemObject") 'file scripting object
    	
    	Set objFolder = objFileScripting.GetFolder(ServerDirectory & ExtraDir & "\" & Folder) 'Remove [ & "\" ] if there is no Extra Directory


	

'________________________________________________
'# Password Protect Folder with PASS on the front

	NoSlash = Replace(Folder, "\","") 'Get rid of slashes so the following code can't be bypassed with more \\\\ slashes	

	PassWordProtect = left(NoSlash, 4)  'PASS
	
		If PassWordProtect = "PASS" then
			Response.Write "Access Denied"
			Response.Flush
			Response.end

			End if
'#_______________________________________________


'mfColor = "000000" 'Main font color

ArrowPic = "../arrow.gif"   'default


'Background colors & Fonts to go with them

		If color = "FFFFFF" then      'white background
			fcolor = "000000" 
			ArrowPic ="../arrow2.gif"     'show black arrow for white background

		elseif color = "003366" then  'drk ice blue
			fcolor = "FFFFFF"

		elseif color = "000000" then  'black background
			fcolor = "FFFFFF"

		elseif color = "CCCCCC" then  'grey background
			fcolor = "000000"
			ArrowPic ="../arrow2.gif"

		elseif color = "660000" then  'maroon
			fcolor = "FFFFFF"

		end if

	
	Response.write "<!--[" & " " & objFolder.Files.Count & " " & "]-->"  'File count

	Response.Write "<body bgcolor=" & chr(34) & "#" & color & chr(34) & " text=" & chr(34) & mfColor & chr(34) & ">"





 	If objFolder.Files.Count <= 0 Then    
       			Response.Write "<font face=" & "Arial color=" & chr(34) & fcolor & chr(34) & "size= 2" & ">" & "No screenshots available for " & "<B>" & Area & "</B>" & "</font>"
  
			ElseIf objFolder.Files.count >= 100 Then
			Response.Write "To many files in folder!"

 			End if

    			Set filecollection = objFolder.Files
    	
    


			For Each filename In filecollection
    			Filename=right(Filename,len(Filename)-InStrRev(Filename, "\"))  'Show images \ urls
    

	

			NoDash = Replace(FileName, "-" , " ")    'File name w/ spaces
			FileDes =  Replace(NoDash, ".gif" , "")  'File Description / no gif
	


'Write pages depending on setting in program:


	If Request.QueryString("settings") = "1" then
		Response.write "<img src=" & chr(34) & ArrowPic  & chr(34) & ">" & "<input type=" & """text""" & " " & "name=" & """textfield""" & " " & "size=" & """70""" & "style=""border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1; background-color: #ffffff" & " " & Chr(34) & "value=" & Chr(34) & strUrlPath & filename & Chr(34) & ">" & "<br>"
		Response.Write "<img src=""" & strUrlPath & filename & """><br>"
		Response.write "<br>"
		Response.write "<br>"
 
	Elseif Request.QueryString("settings") = "2" then

		Response.write "<br>" & ("<font face= Arial color=" & fcolor & " size= 2>")  & FileDes & "</font>" & "</br>"     
		Response.write "<img src=" & chr(34) & ArrowPic  & chr(34) & ">" & "<input type=" & """text""" & " " & "name=" & """textfield""" & " " & "size=" & """70""" & "style=""border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1; background-color: #ffffff" & " " & Chr(34) & "value=" & Chr(34) & strUrlPath & filename & Chr(34) & ">" ' & "<br>"
		Response.write "<br>"
		Response.write "<br>"

	ElseIf Request.QueryString("settings") = "3" then

		Response.write "<br>" & ("<font face= Arial color=" & fcolor & " size= 2>")  & FileDes & "</font>" & "</br>"  
		Response.write "<a href=" & chr(34) & "getimage.asp?img=" & Folder & FileName & "&" & "Information=" & FileDes & chr(34) & ">" & "<img src=" & chr(34) & ArrowPic & chr(34) & "border=" & chr(34) & "0" & chr(34) & ">" & "</a>" & "<input type=" & """text""" & " " & "name=" & """textfield""" & " " & "size=" & """70""" & "style=""border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1; background-color: #ffffff" & " " & Chr(34) & "value=" & Chr(34) & strUrlPath & filename & Chr(34) & ">" & "<br>"
		Response.write "<br>"
		Response.write "<br>"	


	end if


	Next    'display next file with same settings


		'Release them from memory
		Set objFileScripting = Nothing
		Set objFile = Nothing
		Set objFolder = Nothing
		Set Folder = Nothing
		Set Area = Nothing
		Set color = Nothing
		Set PassWordProtect1 = Nothing
		
%>