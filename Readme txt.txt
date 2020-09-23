Read Me
--------------------------------------
ScreenShots by Jason H
irideforlife@irideforlife.com
-Feb 25, 2002
-4:22 AM (haha)
--------------------------------------




1 . What you need to run this: 

IIS 4 or above running or a host with access to your root directory ie (know the location of your folder on the server)  [ c:\Inetpub\wwwroot\ss\ ]
[[ Note: I am not sure if PWS will work, have not tested it yet.]]



###### Files Included: #######

screenshot    	| 	visual basic project
FrmMain	|	main form
FrmGenerate	|	generates html / asp code for your image link

getscreenshot.asp	|	what the vb program uses to call screenshots
getimage.asp		|	to display an image the program uses this page its for settings=3
firstpage.asp		|	for server admin use, so they don't have to edit the program, they just can edit the text in there to hide/show menus etc..

/ss/	(Folder)		|	the default directory of the screenshots and the asp source

Screenshots included = XP / Photoshop / AIM

##########################

Setup: #


1. Put all folders in a folder called "ss" on the www directory also put arrow.gif, arrow2.gif, and firstpage.asp in the ss folder
  >>For example:      c:\inetpub\wwwroot\ss\

If you alread have something named ss in there, then you can change the folder name to whatever you like, you will just have to change the
folder name in the program and in the asp files specified below...


2. Goto the folder "screenshot"   c:\inetpub\wwwroot\ss\screenshot\    and open up getscreenshot in your favorite asp program, and edit the top lines to fit your server:

	ServerDirectory = "c:\inetpub\wwwroot\"    'root directory
  	ServerAddy = "http://localhost/		 'server address 
	ExtraDir = "ss"  		          		 'extra Directories  {change this if you did a different folder for the Screenshots}

Save and close



3. In the folder screenshots

	There is another file called getimage.asp if this: ServerAddy = "http://localhost/ss/"  does NOT  fit how you have this set up then open it and change it.



Admin features:
	There is a feature in this program, that reads firstpage.asp and will hide menus / show menus. It hasn't been tested to a "T" but it works alright. I have commented
it out for now, if you would like to mess with it, un-comment it and then edit the 3 links at the bottom of the page to fit your server / host.




The "Link" in a nut shell ..........

1				2					3						4			5

getscreenshot.asp?		Folder=xp\controlpanel\fonts\	 &	Area=Windows%20XP%20-%20Fonts	&	color=003366	&      settings=1

The asp page			The folder to get				what area you are in to be displayed if no file	background color	what to display / image / text only / text w/ URL



More detail...

3. Area = Area=Windows%20XP%20-%20Fonts  or   Windows XP - Fonts   all this does is show what the page is, so if there is no image available it can display:
No screenshots available for Windows XP - Fonts

4. Background color; which is changed in the program, and the link gets updated to display 003366 background color


5. Image and link  = settings=1
   Text Only	= settings= 2   / low bandwidth servers doesnt display images just file discription and link to it
   text / URL     = settings = 3  / Shows URL and if you click the URL it uses getimage.asp to show the image with file name / discription above it..

--------------------------------------------------------------------


Remember! =  Save your screenshots with "-" for spaces instead of spaces!



Have fun! And happy coding!  If you get a chance vote! If you like my code.. thanks!

Comments, Question, errors? 
irideforlife@irideforlife.com










