VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmGenerate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate Code"
   ClientHeight    =   3450
   ClientLeft      =   7605
   ClientTop       =   6510
   ClientWidth     =   5085
   Icon            =   "frmImg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkGenAuto 
      Caption         =   "Auto Generate"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CheckBox chkAuto 
      Caption         =   "Auto Copy"
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame fHTML 
      Caption         =   "Html / Forum img tags"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4935
      Begin VB.OptionButton html 
         Caption         =   "<img src>"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton html 
         Caption         =   "[img] [/img]"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton html 
         Caption         =   "<img src> <href>"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "The Image Link:"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4935
      Begin VB.TextBox txtLink 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Code:"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4935
      Begin RichTextLib.RichTextBox txtOutput 
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   1931
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmImg.frx":0442
      End
   End
   Begin VB.CommandButton cmdGen2 
      Caption         =   "Generate"
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   3120
      Width           =   735
   End
   Begin VB.Frame fASP 
      Caption         =   "ASP"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      Begin VB.OptionButton asp 
         Caption         =   "Response.write <src img>"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fURL 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtUrl 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   4695
      End
      Begin VB.TextBox txtBorder 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   14
         Text            =   "0"
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "URL"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Border Size"
         Height          =   375
         Left            =   3480
         TabIndex        =   15
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuLang 
      Caption         =   "Language"
      Begin VB.Menu lang 
         Caption         =   "HTML"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu lang 
         Caption         =   "ASP"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmGenerate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub asp_Click(Index As Integer)
    Select Case asp(0)
    
    Case 1
    Call CheckAutoGenerate
    fURL.Visible = False
    frmGenerate.Height = 4110
    cmdCopy.Top = 3120
    cmdGen2.Top = 3120
    chkAuto.Top = 3120
    chkGenAuto.Top = 3120
    End Select
    
    
    

End Sub

Private Sub chkAuto_Click()

        Select Case chkAuto
        
        Case 1
        IChecked = "1"
        
        End Select
        
End Sub

Private Sub cmdCopy_Click()
Clipboard.Clear 'clear contents of clipboard
Clipboard.SetText txtOutput.Text 'set txtOutput to clipboard
End Sub

Private Sub cmdGen2_Click()
Call generateCode 'Call generate code function
End Sub

Private Sub Form_Load()
Dim ClipBoardImg        As String
Dim LinkEnding          As String


    ClipBoardImg = Clipboard.GetText 'put contents of clipboard as a string = ClipBoardImg

    LinkEnding = Right(ClipBoardImg, 3) 'get the last 3 letters of contents in clipboard "gif"

    If LinkEnding = "gif" Then
        txtLink.Text = Clipboard.GetText 'If img link is in clipboard put it in
                                    'txt box
    ElseIf LinkEnding = "jpg" Then
        txtLink.Text = Clipboard.GetText

    End If
    
    If IChecked = "1" Then
    chkAuto.Value = 1
    Else
    End If

End Sub



Private Sub html_Click(Index As Integer)
     
    '# <src img>
    Select Case html(0)
        
        Case 1
        fURL.Visible = False
        frmGenerate.Height = 4110
        cmdCopy.Top = 3120
        cmdGen2.Top = 3120
        chkAuto.Top = 3120
        chkGenAuto.Top = 3120
        Call CheckAutoGenerate
        
        
        
    End Select
        
       
    '# [img]
    Select Case html(1)
        
        Case 1
        fURL.Visible = False
        frmGenerate.Height = 4110
        cmdCopy.Top = 3120
        cmdGen2.Top = 3120
        chkAuto.Top = 3120
        chkGenAuto.Top = 3120
       
       Call CheckAutoGenerate
    
    End Select
    
    
    '# <src img> <href>
    Select Case html(2)
    
        Case 1
        fURL.Visible = True
        frmGenerate.Height = 4665  'Bring form down to show link box
        cmdCopy.Top = 3720
        cmdGen2.Top = 3720
        chkAuto.Top = 3720
        chkGenAuto.Top = 3720
        txtLink.Text = "" ' remove whats there
        
        'Call CheckAutoGenerate   <--- No auto generate because there
                                    ' is still and Xtra field to fill
    
    End Select
    
        
        
        


End Sub

Private Sub lang_Click(Index As Integer)

        Select Case Index

        Case 1
        fHTML.Visible = True
        fASP.Visible = False
        lang(1).Checked = True
        lang(2).Checked = False
        asp(0).Value = False
        
        
        Case 2
        fHTML.Visible = False
        fASP.Visible = True
        lang(1).Checked = False
        lang(2).Checked = True
        html(0).Value = False
        html(1).Value = False
        html(2).Value = False
        
        
End Select


End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub txtBorder_KeyPress(KeyAscii As Integer)
'KeyAscii = IIf(Not KeyAscii = 8 And Not Val((Chr(KeyAscii))) > 0, 0, KeyAscii)
'KeyAscii = IIf(Not KeyAscii = 8 And Not IsNumeric(Chr(KeyAscii)), 0, KeyAscii)
'If Not IsNumeric(txtBorder) Then txtBorder.Text = ""
End Sub


Public Function generateCode()

If txtLink.Text = "" Then: MsgBox "There is no value in the image link!", vbInformation, "Empty Field": Exit Function

txtOutput.Text = "" 'clear before doing the following

If html(0).Value = True Then
txtOutput.Text = "<img src=" & Chr(34) & txtLink.Text & Chr(34) & ">"
ElseIf html(1).Value = True Then
txtOutput.Text = "[img]" & txtLink.Text & "[/img]"
ElseIf html(2).Value = True Then
txtOutput.Text = "<a href=" & Chr(34) & txtUrl.Text & Chr(34) & ">" & "<img src=" & Chr(34) & txtLink.Text & Chr(34) & " " & "border=" & Chr(34) & txtBorder.Text & Chr(34) & "></a>"
ElseIf asp(0).Value = True Then
txtOutput.Text = "<% Response.write " & Chr(34) & "<img src=" & Chr(34) & Chr(34) & txtLink.Text & Chr(34) & Chr(34) & ">" & Chr(34) & "%>"

End If
 '<a href=""><img src="" border=""></a>
 
 'Response.write "<img src=""http://68.6.129.137:81/ss/xp/startbutton/accessories-system-tools.gif"" border=""0""></a>"
'ElseIf asp(1).Value = True Then
'txtOutput.Text = "<% Response.write " & Chr(34) & "<a href=" & Chr(34) & Chr(34) & txtUrl.Text & Chr(34) & Chr(34) & ">" & "<img src=" & Chr(34) & Chr(34) & txtLink.Text & Chr(34) & Chr(34) & " border=" & Chr(34) & Chr(34) & txtBorder.Text & Chr(34) & Chr(34) & "></a>" & Chr(34)




End Function

Private Sub txtLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Show tooltip of whats in there incase its cut off..simply yet useful
txtLink.ToolTipText = txtLink.Text
End Sub

Private Sub txtOutput_Change()
    Select Case chkAuto
    
        Case 1
        Clipboard.Clear 'clear contents of clipboard
        Clipboard.SetText txtOutput.Text 'set txtOutput to clipboard
        Case 0
        
        
        End Select
        


End Sub


Public Function CheckAutoGenerate()
    
    Select Case chkGenAuto

    Case 1
    Call generateCode
    
    Case 0
    End Select
    
 End Function
    
    
