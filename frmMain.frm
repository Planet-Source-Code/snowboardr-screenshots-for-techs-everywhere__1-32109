VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   Caption         =   "[ScreenShots]"
   ClientHeight    =   7455
   ClientLeft      =   3300
   ClientTop       =   4170
   ClientWidth     =   12945
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form10"
   ScaleHeight     =   7455
   ScaleWidth      =   12945
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   3600
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3600
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar PB 
      Height          =   210
      Left            =   7200
      TabIndex        =   6
      Top             =   90
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   0
      OLEDropMode     =   1
   End
   Begin SHDocVwCtl.WebBrowser wb1 
      Height          =   6975
      Left            =   4200
      TabIndex        =   1
      Top             =   360
      Width           =   8775
      ExtentX         =   15478
      ExtentY         =   12303
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin ComctlLib.TreeView TreeView1 
      DragIcon        =   "frmMain.frx":0442
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   12303
      _Version        =   327682
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   0
   End
   Begin VB.Label lbNum 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   12360
      TabIndex        =   8
      Top             =   90
      UseMnemonic     =   0   'False
      Width           =   375
   End
   Begin VB.Label lbPics 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   90
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lbNav 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   Refresh"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   5040
      MouseIcon       =   "frmMain.frx":0594
      TabIndex        =   5
      Top             =   60
      Width           =   855
   End
   Begin VB.Label lbNav 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   Forward"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   6120
      MouseIcon       =   "frmMain.frx":06E6
      TabIndex        =   4
      Top             =   60
      Width           =   855
   End
   Begin VB.Label lbNav 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Back"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4320
      MouseIcon       =   "frmMain.frx":0838
      TabIndex        =   3
      Top             =   60
      Width           =   495
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   1
      Left            =   2280
      Picture         =   "frmMain.frx":098A
      Top             =   2400
      Width           =   240
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   6
      Left            =   2640
      Picture         =   "frmMain.frx":0D32
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   5
      Left            =   2640
      Picture         =   "frmMain.frx":109E
      Top             =   2760
      Width           =   240
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   4
      Left            =   2640
      Picture         =   "frmMain.frx":141C
      Top             =   2400
      Width           =   240
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   3
      Left            =   2280
      Picture         =   "frmMain.frx":17C3
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   2
      Left            =   2280
      Picture         =   "frmMain.frx":1B2F
      Top             =   2760
      Width           =   240
   End
   Begin ComctlLib.ImageList TreeImages 
      Left            =   2280
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Label lbtop 
      BackColor       =   &H00808080&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   20055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuUpload 
         Caption         =   "Upload ScreenShot"
      End
      Begin VB.Menu ln1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu lineo1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHtmlGen 
         Caption         =   "Generate HTML"
      End
      Begin VB.Menu mnuASP 
         Caption         =   "Generate ASP"
      End
      Begin VB.Menu ln3 
         Caption         =   "-"
      End
      Begin VB.Menu BgColor1 
         Caption         =   "Background Color"
         Begin VB.Menu BgColor 
            Caption         =   "White"
            Index           =   1
         End
         Begin VB.Menu BgColor 
            Caption         =   "Black"
            Index           =   2
         End
         Begin VB.Menu BgColor 
            Caption         =   "Grey"
            Index           =   3
         End
         Begin VB.Menu BgColor 
            Caption         =   "Dark ice Blue"
            Index           =   4
         End
         Begin VB.Menu BgColor 
            Caption         =   "Maroon"
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuTop 
         Caption         =   "On Top"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuline 
         Caption         =   "-"
      End
      Begin VB.Menu settings 
         Caption         =   "Images"
         Index           =   1
         Shortcut        =   {F1}
      End
      Begin VB.Menu settings 
         Caption         =   "Text Only"
         Index           =   2
         Shortcut        =   {F2}
      End
      Begin VB.Menu settings 
         Caption         =   "Text / URL"
         Index           =   3
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuline3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFullScreen 
         Caption         =   "Full Screen"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuMini 
         Caption         =   "Minimize"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowSS 
         Caption         =   "Show # of Screenshots"
         Begin VB.Menu mnuSNleft 
            Caption         =   "Left top"
         End
         Begin VB.Menu mnuRightTop 
            Caption         =   "Right top just #"
         End
      End
   End
   Begin VB.Menu mnuHp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuContact 
         Caption         =   "Contact"
         Visible         =   0   'False
         Begin VB.Menu mnuBugs 
            Caption         =   "Report Bugs"
         End
         Begin VB.Menu mnuComments 
            Caption         =   "Comments"
         End
         Begin VB.Menu mnuUpdates 
            Caption         =   "Updates"
         End
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "Update Menus"
      End
      Begin VB.Menu mnuVote 
         Caption         =   "Vote"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
'#By Jason Howard aka snowboardr
'#February 25, 2001
'#
'#If you liked this program, updated it, comments
'#or questions send them to irideforlife@irideforlife.com
'#
'#
'#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#

Public FrontPage  As String
Option Explicit

Dim Forwards As Boolean   'Browser navigation
Dim Backwards As Boolean
Dim nonav As Boolean


Private Enum ObjectType
    otNone = 0
    otprogram = 1 'closed folder
    otGroup = 2
    otsection = 3
    otprogram2 = 4 'opened folder
    otGroup2 = 5
    otsection2 = 6
End Enum

Private SourceNode As Object
Private SourceType As ObjectType
Private TargetNode As Object

Private Function NodeType(mynode As Node) As ObjectType
    If mynode Is Nothing Then
        NodeType = otNone
    Else
        Select Case Left$(mynode.Key, 1)
            Case "f"
                NodeType = otprogram
            Case "g"
                NodeType = otGroup
            Case "p"
                NodeType = otsection
        End Select
    End If
End Function

Private Sub BgColor_Click(Index As Integer)
                                            'Background color
    Select Case Index

        Case 1
            Color = "FFFFFF" '  White background
            Call TreeView1_Click '<-----      If something is selected in treeview it will
                                         '    be clicked again to update background color
            BgColor(1).Checked = True
            BgColor(2).Checked = False
            BgColor(3).Checked = False
            BgColor(4).Checked = False
            BgColor(5).Checked = False

        Case 2
            Color = "000000"  '  Black
            Call TreeView1_Click
            
            BgColor(1).Checked = False
            BgColor(2).Checked = True
            BgColor(3).Checked = False
            BgColor(4).Checked = False
            BgColor(5).Checked = False
            
        Case 3
            Color = "CCCCCC"  '  Grey
            Call TreeView1_Click
            
            BgColor(1).Checked = False
            BgColor(2).Checked = False
            BgColor(3).Checked = True
            BgColor(4).Checked = False
            BgColor(5).Checked = False
        Case 4
            Color = "003366"  '  Dark ice blue
            Call TreeView1_Click
            
            BgColor(1).Checked = False
            BgColor(2).Checked = False
            BgColor(3).Checked = False
            BgColor(4).Checked = True
            BgColor(5).Checked = False
        Case 5
            Color = "660000"  '  Maroon
            Call TreeView1_Click
            
            BgColor(1).Checked = False
            BgColor(2).Checked = False
            BgColor(3).Checked = False
            BgColor(4).Checked = False
            BgColor(5).Checked = True
        End Select
    End Sub


Private Sub Form_Load()
'(must edit link at the bottom to fit your server -------->
'Call GetProgramSettings   Disabled unless you want to use it,its for the Server Admin to edit how the program works via frontpage.asp
Dim L As Integer

'SetNum = "1" 'default setting/ show images/link
Color = "FFFFFF"
'settings(1).Checked = True

FrontPage = "http://localhost/ss/firstpage.asp"  'Front page
'wb1.Navigate FrontPage

wb1.Navigate "http://www.irideforlife.com/psc.asp"

'(must edit link at the bottom to fit your server -------->
'Call GetMainFolder   Disabled unless you want to use it,its for the Server Admin to edit how the program works via frontpage.asp


    ' Load pictures into the Image list.
    For L = 1 To 6
        TreeImages.ListImages.Add , , TreeImage(L).Picture
    Next L
    
    ' Attach the TreeView to the Image list.
    TreeView1.ImageList = TreeImages

    ' Build treeview list
    Call buildTree
    
    If settings(1).Visible = True And settings(2).Visible = True And settings(3) = True Then
    SetNum = "1"
    ElseIf settings(1).Visible = False And settings(2).Visible = True And settings(3) = False Then
    SetNum = "2"
    Else
    End If
    
    

    
End Sub


Private Sub lbNav_Click(Index As Integer)
Select Case Index
    Case 0
        wb1.GoBack
        ChangeColor lbNav(0)  'browser back
    Case 1
        wb1.GoForward
        ChangeColor lbNav(1)  'browser forward
    Case 2
        wb1.Refresh
        ChangeColor lbNav(2)  'browser refresh
End Select
wb1.SetFocus
End Sub

Private Sub lbNav_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ChangeBack lbNav(0)
ChangeBack lbNav(1) 'just messing with color change on a label, nothing big..
ChangeBack lbNav(2)
End Sub

Private Sub lbtop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ChangeBack lbNav(0)
ChangeBack lbNav(1)
ChangeBack lbNav(2)
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuASP_Click()
        
        frmGenerate.Show
        frmGenerate.fHTML.Visible = False 'Hide html options
        frmGenerate.fASP.Visible = True   'Show asp options
        frmGenerate.lang(1).Checked = False  'uncheck html
        frmGenerate.lang(2).Checked = True   'check asp
        frmGenerate.html(0).Value = False
        frmGenerate.html(1).Value = False    'no values to html so It doesnt genrate
        frmGenerate.html(2).Value = False    'html code
        
End Sub
Private Sub mnuCopy_Click()
 wb1.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DONTPROMPTUSER  'copy whats select in browser
End Sub


Private Sub mnuExit_Click()
Dim E As Integer
Dim frm As Form
E% = MsgBox("Are you sure you want to exit?", vbYesNo, "Exit")

If E% = 6 Then '6 = yes

For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next frm
        Else
   
End If
End Sub


Private Sub mnuFullScreen_Click()
    
    If mnuFullScreen.Checked = False Then   'If checked already don't do anything
        frmMain.WindowState = vbMaximized  'FrmMain = full screen
    Else
    End If

End Sub

Private Sub mnuHelp_Click()
MsgBox "View read me", , "Help"
End Sub

Private Sub mnuMini_Click()
frmMain.WindowState = vbMinimized  'Minimize app
End Sub



Private Sub mnuRightTop_Click()
 
 If lbNum.Visible = True Then  ' If its visible then hide it
            mnuRightTop.Checked = False ' un-check menu
            lbNum.Visible = False
            lbPics.Visible = False
        Else
            lbNum.Visible = True      ' If it isn't then show it
            mnuRightTop.Checked = True  ' check menu
            lbPics.Visible = False
        End If
End Sub

Private Sub mnuSNleft_Click()
      If lbPics.Visible = True Then  ' If its visible then hide it
            lbPics.Visible = False
            mnuSNleft.Checked = False ' un-check menu
        Else
            lbPics.Visible = True      ' If it isn't then show it
            mnuSNleft.Checked = True  ' check menu
            lbNum.Visible = False
        End If
End Sub

Private Sub mnuTop_Click()
If mnuTop.Checked = False Then
    MakeTopMost Me.hwnd 'Keep on top
    mnuTop.Checked = True 'check menu
    frmMain.Caption = "[ScreenShots] - On top"  'set form caption
 Else
    MakeNormal Me.hwnd 'make normal
    mnuTop.Checked = False 'un-check
    frmMain.Caption = "[ScreenShots]"  'set form caption back
End If
End Sub


Public Function buildTree()

'Build our big treeview..

Dim program As Node
Dim section As Node
Dim group As Node

'# Windows XP
Set program = TreeView1.Nodes.Add(, , "f Windows XP", "Windows XP", otprogram, otprogram2)
    Set group = TreeView1.Nodes.Add(program, tvwChild, "g Start button", "Start button", otGroup, otGroup2)
    Set group = TreeView1.Nodes.Add(program, tvwChild, "g Control Panel", "Control Panel", otGroup, otGroup2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p add hardware", "Add Hardware", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p add or remove", "Add or Remove", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p admin tools", "Admin Tools", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p display", "Display", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p folder options", "Folder options", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p fonts", "Fonts", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p internet options", "Internet Options", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p mouse", "Mouse", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p network connections", "Network connections", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p phone and modem", "Phone and modem", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p power options", "Power options", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p printers and faxes", "Printers and faxes", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p regional and language", "Regional and language", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p scanners and cameras", "Scanners and cameras", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p scheduled tasks", "scheduled tasks", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p sounds and audio", "Sounds and audio", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p speech", "Speech", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p system", "System", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p taskbar and menu", "Taskbar and menu", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p user accounts", "User accounts", otsection, otsection2)
    Set group = TreeView1.Nodes.Add(program, tvwChild, "g My Computer", "My Computer", otGroup, otGroup2)
    Set group = TreeView1.Nodes.Add(program, tvwChild, "g IIS 5", "IIS 5", otGroup, otGroup2)
    Set group = TreeView1.Nodes.Add(program, tvwChild, "g gpedit.msc", "gpedit.msc", otGroup, otGroup2)

'# Windows 2000
    Set program = TreeView1.Nodes.Add(, , "f Windows 2000", "Windows 2000", otprogram, otprogram2)
    Set group = TreeView1.Nodes.Add(program, tvwChild, "g Start button 2k", "Start button", otGroup, otGroup2)
    Set group = TreeView1.Nodes.Add(program, tvwChild, "g Control Panel 2k", "Control Panel", otGroup, otGroup2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p add hardware 2k", "Add Hardware", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p add or remove 2k", "Add or Remove", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p admin tools 2k", "Admin Tools", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p display 2k", "Display", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p folder options 2k", "Folder options", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p fonts 2k", "Fonts", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p internet options 2k", "Internet Options", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p mouse 2k", "Mouse", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p network connections 2k", "Network connections", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p phone and modem 2k", "Phone and modem", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p power options 2k", "Power options", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p printers and faxes 2k", "Printers and faxes", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p regional and language 2k", "Regional and language", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p scanners and cameras 2k", "Scanners and cameras", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p scheduled tasks 2k", "scheduled tasks", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p sounds and audio 2k", "Sounds and audio", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p speech 2k", "Speech", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p system 2k", "System", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p taskbar and menu 2k", "Taskbar and menu", otsection, otsection2)
        Set section = TreeView1.Nodes.Add(group, tvwChild, "p user accounts 2k", "User accounts", otsection, otsection2)
    Set group = TreeView1.Nodes.Add(program, tvwChild, "g My Computer 2k", "My Computer", otGroup, otGroup2)
    Set group = TreeView1.Nodes.Add(program, tvwChild, "g IIS 5 2k", "IIS 5", otGroup, otGroup2)
    Set group = TreeView1.Nodes.Add(program, tvwChild, "g gpedit.msc 2k", "gpedit.msc", otGroup, otGroup2)
    Set group = TreeView1.Nodes.Add(program, tvwChild, "g Administrative Tools 2k", "Administrative Tools", otGroup, otGroup2)
    Set group = TreeView1.Nodes.Add(program, tvwChild, "g Printers and faxes 2k", "Printers and faxes", otGroup, otGroup2)
  
'#Windows 98
'modem
    Set program = TreeView1.Nodes.Add(, , "f Windows 98", "Windows 98", otprogram, otprogram2)
        Set group = TreeView1.Nodes.Add(program, tvwChild, "g Start button 98", "Start button", otGroup, otGroup2)
        Set group = TreeView1.Nodes.Add(program, tvwChild, "g Control Panel 98", "Control Panel", otGroup, otGroup2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p accessiblity opt", "Accessibility Options", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p add hardware 98", "Add Hardware", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p add or remove 98", "Add or Remove", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p date-time", "Date / Time", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p display 98", "Display", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p fonts 98", "Fonts", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p internet options 98", "Internet Options", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p mouse 98", "Mouse", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p network connections 98", "Network connections", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p mail", "Mail", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p modem", "Modem", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p power management, 98", "Power management,", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p printers and faxes 98", "Printers and faxes", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p regional and language 98", "Regional and language", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p scheduled tasks 98", "scheduled tasks", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p sounds", "Sounds and audio", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p system 98", "System", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p keyboard", "Keyboard", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "p taskbar and menu 98", "Taskbar and menu", otsection, otsection2)
            Set section = TreeView1.Nodes.Add(group, tvwChild, "users", "Users", otsection, otsection2)

        Set group = TreeView1.Nodes.Add(program, tvwChild, "g My Computer 98", "My Computer", otGroup, otGroup2)
        Set group = TreeView1.Nodes.Add(program, tvwChild, "g PWS 98", "PWS", otGroup, otGroup2)
        Set group = TreeView1.Nodes.Add(program, tvwChild, "g Printers and faxes 98", "Printers and faxes", otGroup, otGroup2)

'#Internet explorer
    Set program = TreeView1.Nodes.Add(, , "f Internet Explorer", "Internet Explorer", otprogram, otprogram2)
        Set group = TreeView1.Nodes.Add(program, tvwChild, "g Main ie", "Main", otGroup, otGroup2)
        Set group = TreeView1.Nodes.Add(program, tvwChild, "g Menus ie", "Menus", otGroup, otGroup2)
        Set group = TreeView1.Nodes.Add(program, tvwChild, "g Toolbar ie", "Toolbar", otGroup, otGroup2)
        Set group = TreeView1.Nodes.Add(program, tvwChild, "g Tools > Internet > Options", "Tools > Internet > Options", otGroup, otGroup2)
'#AIM
    Set program = TreeView1.Nodes.Add(, , "f AIM", "AIM", otprogram, otprogram2)
        Set group = TreeView1.Nodes.Add(program, tvwChild, "g Main windows AIM", "Main Windows", otGroup, otGroup2)
        Set group = TreeView1.Nodes.Add(program, tvwChild, "g My AIM > Edit Opt > Edit Prefs", "My AIM > Edit Opt > Edit Prefs", otGroup, otGroup2)
        Set group = TreeView1.Nodes.Add(program, tvwChild, "g Menus AIM", "Menus", otGroup, otGroup2)
'#MSN
    Set program = TreeView1.Nodes.Add(, , "f MSN", "MSN", otprogram, otprogram2)
        Set group = TreeView1.Nodes.Add(program, tvwChild, "g Tools > Options", "Tools > Options", otGroup, otGroup2)
        Set group = TreeView1.Nodes.Add(program, tvwChild, "g Menus MSN", "Menus", otGroup, otGroup2)

'# MS Word
    Set program = TreeView1.Nodes.Add(, , "f Microsoft Word", "Microsoft Word XP", otprogram, otprogram2)
            Set group = TreeView1.Nodes.Add(program, tvwChild, "g windows wxp", "Main Windows", otGroup, otGroup2)
            Set group = TreeView1.Nodes.Add(program, tvwChild, "g options wxp", "Tools > Options", otGroup, otGroup2)
            Set group = TreeView1.Nodes.Add(program, tvwChild, "g menu wxp", "Menus", otGroup, otGroup2)
            Set group = TreeView1.Nodes.Add(program, tvwChild, "g toolbar wxp", "Toolbar", otGroup, otGroup2)
            Set group = TreeView1.Nodes.Add(program, tvwChild, "g Printing wxp", "Printing", otGroup, otGroup2)
            Set group = TreeView1.Nodes.Add(program, tvwChild, "g wordart wxp", "Word Art", otGroup, otGroup2)


'# To be continued....

End Function



Private Sub mnuUpdate_Click()

Call GetProgramSettings
Call GetProgramSettings1
Call GetProgramSettings2

End Sub

Private Sub mnuUpload_Click()
MsgBox "Will include in next Version!", , "Upload image"
End Sub

Private Sub mnuVote_Click()
wb1.Navigate "http://www.irideforlife.com/psc.asp"
Me.WindowState = vbMaximized

End Sub

Private Sub TreeView1_Click()

'Call GotoFolder (" a ", " b" )
'a = Folder to look in [string]
'b = What it is, so if there is nothing there its easy to display... [string]

On Error GoTo err91

'# Windows Xp
If TreeView1.SelectedItem.Key = "g Start button" Then
        'Call function in mConst module
        
         Call GotoFolder("xp\startbutton\", "Windows 2000 - Start Button")


'Control panel section >
ElseIf TreeView1.SelectedItem.Key = "p add hardware" Then
       Call GotoFolder("xp\controlpanel\addhardware\", "Windows XP - Add Hardware")
        

ElseIf TreeView1.SelectedItem.Key = "p add or remove" Then
        Call GotoFolder("xp\controlpanel\addorremove\", "Windows XP - Add or remove")
   
ElseIf TreeView1.SelectedItem.Key = "p admin tools" Then
        Call GotoFolder("xp\controlpanel\admintools\", "Windows XP - Admin tools")
   
ElseIf TreeView1.SelectedItem.Key = "p display" Then
        Call GotoFolder("xp\controlpanel\display\", "Windows XP - Display")

ElseIf TreeView1.SelectedItem.Key = "p folder options" Then
        Call GotoFolder("xp\controlpanel\folderoptions\", "Windows XP - Folder Options")

ElseIf TreeView1.SelectedItem.Key = "p fonts" Then
        Call GotoFolder("xp\controlpanel\fonts\", "Windows XP - Fonts")

ElseIf TreeView1.SelectedItem.Key = "p internet options" Then
     Call GotoFolder("xp\controlpanel\internetoptions\", "Windows XP - Internet options")

ElseIf TreeView1.SelectedItem.Key = "p mouse" Then
     Call GotoFolder("xp\controlpanel\mouse\", "Windows XP - Mouse ")

ElseIf TreeView1.SelectedItem.Key = "p network connections" Then
     Call GotoFolder("xp\controlpanel\networkconnections\", "Windows XP - Network Connections")

ElseIf TreeView1.SelectedItem.Key = "p phone and modem" Then
     Call GotoFolder("xp\controlpanel\phoneandmodem\", "Windows XP - Phone and modem")

ElseIf TreeView1.SelectedItem.Key = "p power options" Then
     Call GotoFolder("xp\controlpanel\poweroptions\", "Windows XP - Power Options")

ElseIf TreeView1.SelectedItem.Key = "p printers and faxes" Then
     Call GotoFolder("xp\controlpanel\printersandfaxes\", "Windows XP - Printers and faxes")

ElseIf TreeView1.SelectedItem.Key = "p regional and language" Then
     Call GotoFolder("xp\controlpanel\regionalandlanguage\", "Windows XP - Regional Language")

ElseIf TreeView1.SelectedItem.Key = "p scanners and cameras" Then
     Call GotoFolder("xp\controlpanel\scannersandcameras\", "Windows XP - Scanners and cameras")

ElseIf TreeView1.SelectedItem.Key = "p scheduled tasks" Then
     Call GotoFolder("xp\controlpanel\scheduledtasks\", "Windows XP - Scheduled tasks")

ElseIf TreeView1.SelectedItem.Key = "p sounds and audio" Then
     Call GotoFolder("xp\controlpanel\soundsandaudio\", "Windows XP - Sounds and audio ")

ElseIf TreeView1.SelectedItem.Key = "p speech" Then
     Call GotoFolder("xp\controlpanel\speech\", "Windows XP - Speech")

ElseIf TreeView1.SelectedItem.Key = "p system" Then
     Call GotoFolder("xp\controlpanel\system\", "Windows XP - System ")

ElseIf TreeView1.SelectedItem.Key = "p taskbar and menu" Then
     Call GotoFolder("xp\controlpanel\taskbarandmenu\", "Windows XP -  Taskbar and menu")

ElseIf TreeView1.SelectedItem.Key = "p user accounts" Then
     Call GotoFolder("xp\controlpanel\useraccounts\", "Windows XP - User Accounts")

'< End Control panel section [ X P ]

ElseIf TreeView1.SelectedItem.Key = "g My Computer" Then
     Call GotoFolder("xp\mycomputer\", "Windows XP - My Computer")

ElseIf TreeView1.SelectedItem.Key = "g IIS 5" Then
     Call GotoFolder("xp\iis\", "Windows XP - IIS")

ElseIf TreeView1.SelectedItem.Key = "g gpedit.msc" Then
    Call GotoFolder("xp\gpedit\", "Windows XP - gpedit.msc")

ElseIf TreeView1.SelectedItem.Key = "g Administrative Tools" Then
    Call GotoFolder("xp\controlpanel\admintools\", "Windows XP - Admin tools")

'< End XP

'# Windows 2000 pro

ElseIf TreeView1.SelectedItem.Key = "g Start button 2k" Then
    Call GotoFolder("windows2000\startbutton\", "Windows 2000 -  Start button")

ElseIf TreeView1.SelectedItem.Key = "g My Computer 2k" Then
    Call GotoFolder("windows2000\gpedit\", "Windows 2000 - My Computer")

ElseIf TreeView1.SelectedItem.Key = "g IIS 5 2k" Then
    Call GotoFolder("windows2000\iis\", "Windows 2000 - IIS ")

ElseIf TreeView1.SelectedItem.Key = "g gpedit.msc 2k" Then
    Call GotoFolder("windows2000\gpedit\", "Windows 2000 - gpedit.msc")

ElseIf TreeView1.SelectedItem.Key = "g Administrative Tools 2k" Then
    Call GotoFolder("windows2000\controlpanel\admintools\", "Windows 2000 - Admin Tools ")
    
'## Windows 2000 Control panel >>

ElseIf TreeView1.SelectedItem.Key = "p add hardware 2k" Then
    Call GotoFolder("windows2000\controlpanel\addhardware\", "Windows 2000 -  Add Hardware")

ElseIf TreeView1.SelectedItem.Key = "p add or remove 2k" Then
    Call GotoFolder("windows2000\controlpanel\addorremove\", "Windows 2000 - Add or remove ")

ElseIf TreeView1.SelectedItem.Key = "p admin tools 2k" Then
    Call GotoFolder("windows2000\controlpanel\admintools\", "Windows 2000 -  Admin tools")

ElseIf TreeView1.SelectedItem.Key = "p display 2k" Then
    Call GotoFolder("windows2000\controlpanel\display\", "Windows 2000 - Display")

ElseIf TreeView1.SelectedItem.Key = "p folder options 2k" Then
    Call GotoFolder("windows2000\controlpanel\folderoptions\", "Windows 2000 - Folder Options")

ElseIf TreeView1.SelectedItem.Key = "p fonts 2k" Then
    Call GotoFolder("windows2000\controlpanel\fonts\", "Windows 2000 - Fonts")

ElseIf TreeView1.SelectedItem.Key = "p internet options 2k" Then
    Call GotoFolder("windows2000\controlpanel\internetoptions\", "Windows 2000 - Internet Options")

ElseIf TreeView1.SelectedItem.Key = "p mouse 2k" Then
    Call GotoFolder("windows2000\controlpanel\mouse\", "Windows 2000 - Mouse")

ElseIf TreeView1.SelectedItem.Key = "p network connections 2k" Then
    Call GotoFolder("windows2000\controlpanel\networkconnections\", "Windows 2000 - Network Connections ")

ElseIf TreeView1.SelectedItem.Key = "p phone and modem 2k" Then
    Call GotoFolder("windows2000\controlpanel\phoneandmodem\", "Windows 2000 - Phone and modem")

ElseIf TreeView1.SelectedItem.Key = "p power options 2k" Then
    Call GotoFolder("windows2000\controlpanel\poweroptions\", "Windows 2000 - Printers and Faxes")

ElseIf TreeView1.SelectedItem.Key = "p printers and faxes 2k" Then
    Call GotoFolder("windows2000\controlpanel\printersandfaxes\", "Windows 2000 - Printers and Faxes")

ElseIf TreeView1.SelectedItem.Key = "p regional and language 2k" Then
    Call GotoFolder("windows2000\controlpanel\regionalandlanguage\", "Windows 2000 - Regional Languages")

ElseIf TreeView1.SelectedItem.Key = "p scanners and cameras 2k" Then
    Call GotoFolder("windows2000\controlpanel\scannersandcameras\", "Windows 2000 - Scanners and cameras")

ElseIf TreeView1.SelectedItem.Key = "p scheduled tasks 2k" Then
   Call GotoFolder("windows2000\controlpanel\scheduledtasks\", "Windows 2000 - Scheduled tasks")

ElseIf TreeView1.SelectedItem.Key = "p sounds and audio 2k" Then
    Call GotoFolder("windows2000\controlpanel\soundsandaudio\", "Windows 2000 - Sound and audio")

ElseIf TreeView1.SelectedItem.Key = "p speech 2k" Then
    Call GotoFolder("windows2000\controlpanel\speech\", "Windows 2000 - Speech ")

ElseIf TreeView1.SelectedItem.Key = "p system 2k" Then
    Call GotoFolder("windows2000\controlpanel\system\", "Windows 2000 - System")

ElseIf TreeView1.SelectedItem.Key = "p taskbar and menu 2k" Then
    Call GotoFolder("windows2000\controlpanel\taskbarandmenu\", "Windows 2000 - Taskbar and menu")

ElseIf TreeView1.SelectedItem.Key = "p user accounts 2k" Then
    Call GotoFolder("windows2000\controlpanel\useraccounts\", "Windows 2000 - User Accounts")

'# Windows 98
ElseIf TreeView1.SelectedItem.Key = "g Start button 98" Then
    Call GotoFolder("windows98\startbutton\", "Windows 98 - Start Button ")

ElseIf TreeView1.SelectedItem.Key = "g PWS 98" Then
    Call GotoFolder("windows98\pws\", "Windows 98 - PWS")
    '# Windows 98 Control panel >>>>>>>>>

ElseIf TreeView1.SelectedItem.Key = "p add hardware 98" Then
    Call GotoFolder("windows98\controlpanel\addhardware\", "Windows 98 - Add hardware")

ElseIf TreeView1.SelectedItem.Key = "p add or remove 98" Then
    Call GotoFolder("windows98\controlpanel\addorremove\", "Windows 98 - Add or remove")

ElseIf TreeView1.SelectedItem.Key = "p display 98" Then
    Call GotoFolder("windows98\controlpanel\display\", "Windows 98 - Display")

ElseIf TreeView1.SelectedItem.Key = "p folder options 98" Then
    Call GotoFolder("windows98\controlpanel\folderoptions\", "Windows 98 - Folder Options")

ElseIf TreeView1.SelectedItem.Key = "p fonts 98" Then
    Call GotoFolder("windows98\controlpanel\fonts\", "Windows 98 - Fonts")

ElseIf TreeView1.SelectedItem.Key = "p internet options 98" Then
    Call GotoFolder("windows98\controlpanel\internetoptions\", "Windows 98 - Internet Options ")

ElseIf TreeView1.SelectedItem.Key = "p mouse 98" Then
    Call GotoFolder("windows98\controlpanel\mouse\", "Windows 98 - Mouse ")

ElseIf TreeView1.SelectedItem.Key = "p network connections 98" Then
    Call GotoFolder("windows98\controlpanel\networkconnections\", "Windows 98 - Network Connections")

ElseIf TreeView1.SelectedItem.Key = "p date-time" Then
    Call GotoFolder("windows98\controlpanel\date-time\", "Windows 98 - Phone and modem")

ElseIf TreeView1.SelectedItem.Key = "mail" Then
    Call GotoFolder("windows98\controlpanel\mail\", "Windows 98 - Mail")

ElseIf TreeView1.SelectedItem.Key = "p printers and faxes 98" Then
    Call GotoFolder("windows98\controlpanel\printersandfaxes\", "Windows 98 - Printers and faxes")

ElseIf TreeView1.SelectedItem.Key = "p regional and language 98" Then
    Call GotoFolder("windows98\controlpanel\regionalandlanguage\", "Windows 98 - Regional and Languages")

ElseIf TreeView1.SelectedItem.Key = "p Keyboard" Then
    Call GotoFolder("windows98\controlpanel\keyboard\", "Windows 98 - Keyboard")

ElseIf TreeView1.SelectedItem.Key = "p scheduled tasks" Then
    Call GotoFolder("windows98\controlpanel\scheduledtasks\", "Windows 98 -  Scheduled tasks")

ElseIf TreeView1.SelectedItem.Key = "p sounds" Then
    Call GotoFolder("windows98\controlpanel\sounds\", "Windows 98 - Sounds")

ElseIf TreeView1.SelectedItem.Key = "p system 98" Then
    Call GotoFolder("windows98\controlpanel\system\", "Windows 98 - System")

ElseIf TreeView1.SelectedItem.Key = "p taskbar and menu 98" Then
   Call GotoFolder("windows98\controlpanel\taskbarandmenu\", "Windows 98 - Taskbar and menu")

ElseIf TreeView1.SelectedItem.Key = "users" Then
   Call GotoFolder("windows98\controlpanel\users\", "Windows 98 - Users")

'# End windows 98 control panel #

ElseIf TreeView1.SelectedItem.Key = "g My Computer 98" Then
    Call GotoFolder("windows98\mycomputer\", "Windows 98 - My Computer")

ElseIf TreeView1.SelectedItem.Key = "g Printers and faxes 98" Then
    Call GotoFolder("windows98\controlpanel\printersandfaxes\", "Windows 98 - Printers and faxes")

'IE Section >
ElseIf TreeView1.SelectedItem.Key = "g Menus ie" Then
    Call GotoFolder("ie\menu\", "Internet Explorer - Menus")

ElseIf TreeView1.SelectedItem.Key = "g Toolbar ie" Then
    Call GotoFolder("ie\toolbar\", "Internet Explorer - Toolbar")

ElseIf TreeView1.SelectedItem.Key = "g Tools > Internet > Options" Then
    Call GotoFolder("ie\options\", "Internet Explorer - Tools > Internet > Options")

ElseIf TreeView1.SelectedItem.Key = "g Main ie" Then
    Call GotoFolder("ie\main\", "Internet Explorer - Main")

'aim section >
ElseIf TreeView1.SelectedItem.Key = "g Main windows AIM" Then
    Call GotoFolder("AIM\main\", "AIM - Main")

ElseIf TreeView1.SelectedItem.Key = "g My AIM > Edit Opt > Edit Prefs" Then
    Call GotoFolder("AIM\edit\", "AIM - Edit")

ElseIf TreeView1.SelectedItem.Key = "g Menus AIM" Then
    Call GotoFolder("AIM\menu\", "AIM - Menus")
'< end aim section

'MSN section >
ElseIf TreeView1.SelectedItem.Key = "g Tools > Options" Then
    Call GotoFolder("MSN\edit\", "MSN - Tools > Options")

ElseIf TreeView1.SelectedItem.Key = "g Menus MSN" Then
    Call GotoFolder("MSN\menu\", "AIM - Menus")

'End If '< end MSN section

ElseIf TreeView1.SelectedItem.Key = "g windows wxp" Then
    Call GotoFolder("mswordxp\windows\", "MS Word XP - Main windows")


ElseIf TreeView1.SelectedItem.Key = "g options wxp" Then
    Call GotoFolder("mswordxp\options\", "MS Word XP - Tools > Options")
    
    
ElseIf TreeView1.SelectedItem.Key = "g menu wxp" Then
    Call GotoFolder("mswordxp\menu\", "MS Word XP - Menus")
    
    
ElseIf TreeView1.SelectedItem.Key = "g toolbar wxp" Then
    Call GotoFolder("mswordxp\toolbar\", "MS Word XP - Toolbar")
    
ElseIf TreeView1.SelectedItem.Key = "g Printing wxp" Then
    Call GotoFolder("mswordxp\printing\", "MS Word XP - Printing")
    
    
ElseIf TreeView1.SelectedItem.Key = "g wordart wxp" Then
    Call GotoFolder("mswordxp\wordart\", "MS Word XP - WordArt")

err91:
If Err.Number = 91 Then
End If
End If

End Sub

Private Sub Form_Resize()
'# Thanks to wastingtape for help on this resize code
'# ---------------------------------------------------

     On Error GoTo err380
   If Not Me.WindowState = vbMinimized Then  'If minimized don't do the following:
        TreeView1.Width = (Me.Width * 0.28)
        
        wb1.Left = TreeView1.Left + TreeView1.Width + 70
        wb1.Width = Me.Width - TreeView1.Width - 200    'Make controls fit window
        wb1.Height = Me.Height - 1070                    'when Resized.
        TreeView1.Height = wb1.Height                   'Bit messy but it works.
   End If
   
  
err380:
  If Err.Number = 380 Then   'when you bring the right/bottom up to the left top it gave
                             '380 error this traps it
  End If
    
    
    If lbPics.Visible = True Then  ' I try and try to get lbnum to not show but it does any ways
    lbNum.Visible = False
    Else
    lbNum.Left = Me.Width - lbNum.Width
    End If
    
    

  
  
    If Me.Width <= 7500 Then  'If width of the form gets to a navigation button then hide it
        lbNum.Visible = False   'which is at 7250 width..
    Else
        lbNum.Visible = True    'Else show it..
    End If
  
  
  
    If frmMain.WindowState = vbMaximized Then  'Check full screen in menu if it is
        mnuFullScreen.Checked = True
        mnuFullScreen.Enabled = False
    Else
        mnuFullScreen.Checked = False
        mnuFullScreen.Enabled = True
    End If
  
End Sub

Private Sub settings_Click(Index As Integer)
Dim NewSetting              As String

Call removeSetting 'remove setting from current link
                    ' so we can add it again depending on which case is selected

        Select Case Index
             
        Case 1
           
            NewSetting = NewUrl & "settings=1" 'change setting on url
            wb1.Navigate NewSetting  'Goto newly created URL
            
            settings(1).Checked = True
            settings(2).Checked = False
            settings(3).Checked = False
           
            SetNum = "1"  ' setting number
            
        Case 2
            NewSetting = NewUrl & "settings=2"
            wb1.Navigate NewSetting
          
            settings(2).Checked = True
            settings(3).Checked = False
            settings(1).Checked = False
            
            SetNum = "2"
        
        Case 3
             
             NewSetting = NewUrl + "settings=3"
             wb1.Navigate NewSetting
           
             settings(3).Checked = True
             settings(2).Checked = False
             settings(1).Checked = False
            
             SetNum = "3"
           
        
End Select

End Sub


Private Sub wb1_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)

On Error Resume Next                  ' wb1 will error if allowed to go backwards with
        Select Case Command           ' no backwards to go..
            Case CSC_NAVIGATEFORWARD  'If allowed to go forwards enable the label
            Forwards = Enable
        
        Case CSC_NAVIGATEBACK
            Backwards = Enable        'If allowed to go backwards enable the label
        
        End Select


        lbNav(0).Enabled = Backwards
        lbNav(1).Enabled = Forwards
        
        If Command = -1 Then Exit Sub

End Sub

Private Sub wb1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Call GetScreenShotNumber  'When download of page is complete call this function
End Sub

Private Sub wb1_DownloadComplete()

currentURL = wb1.LocationURL  'When browser finishes its navigation it saves link
                              'as string to currentURL

End Sub

Private Sub wb1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
    If PB.Value = 100 Then   'If progress bar reaches 100 hide it
        PB.Visible = False
    Else
        PB.Visible = True     'If its not at 100 then keep it visible
    End If
    If Progress = -1 Then PB.Value = 100
    If Progress > 0 And ProgressMax > 0 Then
        PB.Value = Progress * 100 / ProgressMax
    End If
     
End Sub


Public Function removeSetting() 'take off setting so its easy to add it again

Dim CurrentSetting

CurrentSetting = Right(currentURL, 1)  'Get end of currentURL to know what the setting is.
If CurrentSetting = "1" Then
    NewUrl = Replace(currentURL, "settings=1", "")  'replace setting=1 with nothing

ElseIf CurrentSetting = "2" Then
    NewUrl = Replace(currentURL, "settings=2", "")

ElseIf CurrentSetting = "3" Then
    NewUrl = Replace(currentURL, "settings=3", "")
End If

End Function

Public Function GetProgramSettings()
Dim source As String
Dim BackGround As Boolean
 

        '   Get source of settings page
        source = Inet1.OpenURL("http://localhost/ss/firstpage.asp")
        '   see if the background menu is suppose to be shown
        BackGround = InStr(source, "background:on")
        '   if its other then "on" then menu will be hidden from user
    
    If BackGround = False Then
        BgColor1.Visible = False
        ln3.Visible = False
        mnuUpdate.Visible = True
    Else
        BgColor1.Visible = True  'So if its already gone, see if the settings have been updated.
        ln3.Visible = True
        mnuUpdate.Visible = False
    End If
   
   
Call GetProgramSettings1  'At the end so the inet control isnt used more then once at one time
                            
End Function


Public Function GetProgramSettings1()
Dim source1 As String
Dim img As Boolean
        source1 = Inet1.OpenURL("http://localhost/ss/firstpage.asp")
        
        img = InStr(source1, "image:on")
        
        If img = False Then
            settings(1).Visible = False
            mnuUpdate.Visible = True
        
        Else
            settings(1).Visible = True
        End If
        
Call GetProgramSettings2

End Function

Public Function GetProgramSettings2()
Dim source2 As String
Dim textURL As Boolean

        source2 = Inet1.OpenURL("http://localhost/ss/firstpage.asp")
        
        textURL = InStr(source2, "textURL:on")
        
        If textURL = False Then
            settings(3).Visible = False
            mnuUpdate.Visible = True
        Else
            settings(3).Visible = True
        End If

End Function

Public Function GetScreenShotNumber()
Dim fSource As String
Dim sNumb() As String
Dim CurrentSetting As String


On Error GoTo NoFileNumber
CurrentSetting = Right(currentURL, 1)   'Get current setting

'If there is no number at the end then
'don't read the page cause we know its not the right page
'If CurrentSetting = 1 Or 2 Or 3 Then
                                      
If IsNumeric(CurrentSetting) Then  'If setting = a number then continue:

    fSource = Inet2.OpenURL(currentURL)
    
    sNumb() = Split(fSource, " ", 8) 'Returns number of files on page which is outputed in each page
                                        'like this <!--[ xxx ]--> where xxx is ammount of files
    
 Call PrintImageCount(sNumb(1))    'Call function which then shows our number to the label

End If
 


               
NoFileNumber: If Err.Number = 9 Then
End If

End Function

Public Function GetMainFolder()
Dim MainFolderSource As String
Dim mfn() As String
'http://68.6.129.137:81/ss/firstpage.asp

'Get main server folder from text..
MainFolderSource = Inet2.OpenURL("http://localhost/ss/firstpage.asp")
mfn() = Split(MainFolderSource, "%", 8)

    If mfn(1) = "" Then
        MainFolder = "ss"
        Else
    MainFolder = mfn(1)
    End If


End Function


Public Function PrintImageCount(Count As String)

If Count <= 100 Then
lbPics.Caption = "Screen shots available: " & Count
lbNum.Caption = Count
lbNum.ToolTipText = "Screen shots on page [" & Count & "]"
Else
lbPics.Caption = "Screen shots available: -"
lbNum.Caption = "-"
End If

End Function
