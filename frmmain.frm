VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Simple Web Server"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Blocked IP's..."
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Full Client Log       (From Current to Past)"
      Height          =   4695
      Left            =   4320
      TabIndex        =   10
      Top             =   240
      Width           =   3975
      Begin VB.TextBox Text2 
         Height          =   4095
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Custom 404 message..."
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Activity Log"
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   3975
      Begin VB.ListBox List1 
         Height          =   1035
         ItemData        =   "frmmain.frx":0000
         Left            =   120
         List            =   "frmmain.frx":0002
         TabIndex        =   8
         Top             =   240
         Width           =   3735
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmmain.frx":0004
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   960
      Top             =   2280
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3120
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton Command4 
         Caption         =   "New HTML"
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   2880
         Width           =   975
      End
      Begin VB.FileListBox File1 
         Height          =   1455
         Left            =   1440
         Pattern         =   "*.html;*.jpg;*.htm;*.gif"
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   975
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3413
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Files (Pages + Images)"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   350
         Left            =   1440
         TabIndex        =   3
         Text            =   "C:\"
         Top             =   360
         Width           =   2295
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   3120
         Top             =   2040
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   2
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmmain.frx":00D9
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmmain.frx":03F3
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Home Directory"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   440
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dImage As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Dim sResponse As String
Dim HomePageFile As String
Dim msg404 As String
Dim BlockedIPs As String



Private Sub Command1_Click()
Text1_Change

End Sub

Private Sub Command2_Click()
msg = InputBox("404 message HTML:")
Open "C:\404.html" For Output As #1
Print #1, msg
Close #1
msg404 = msg

End Sub

Private Sub Command3_Click()
On Error GoTo NewList
Open "C:\BlockedIPs.txt" For Input As #1
 Do While Not EOF(1)
 Input #1, BlockedIPs
 Loop
Close #1

BlockedIPs = InputBox("Blocked IP's (space to seperate)", "Blocked IP Addresses", BlockedIPs)
Open "C:\blockedips.txt" For Output As #1
Print #1, BlockedIPs
Close #1
Exit Sub
NewList:
BlockedIPs = InputBox("Blocked IP's (Space to seperate)", "Blocked IP Addresses", "")
Open "C:\blockedips.txt" For Output As #1
Print #1, BlockedIPs
Close #1

End Sub

Private Sub Command4_Click()
pagename = InputBox("File Name:")
Open Text1.Text & pagename For Output As #1
Print #1, ""
Close #1
File1.Refresh
ListView1.ListItems.Clear

filecount = File1.ListCount

For i = 1 To filecount - 1
  If UCase(Right(File1.List(i), 1)) = "L" Or UCase(Right(File1.List(i), 1)) = "M" Then ListView1.ListItems.Add i, "", File1.List(i), , 1
  If UCase(Right(File1.List(i), 1)) = "F" Or UCase(Right(File1.List(i), 1)) = "G" Then ListView1.ListItems.Add i, "", File1.List(i), , 2
  
  
 Next i
 
End Sub

Private Sub Form_Load()
Winsock1.LocalPort = 80
Winsock1.Listen
HomePageFile = Text1.Text & "index.html"
File1.Path = Text1.Text
filecount = File1.ListCount






For i = 1 To filecount - 1
  If UCase(Right(File1.List(i), 1)) = "L" Or UCase(Right(File1.List(i), 1)) = "M" Then ListView1.ListItems.Add i, "", File1.List(i), , 1
  If UCase(Right(File1.List(i), 1)) = "F" Or UCase(Right(File1.List(i), 1)) = "G" Then ListView1.ListItems.Add i, "", File1.List(i), , 2
  
  
 Next i
 On Error GoTo new404
 Open "C:\404.html" For Input As #1
 Do While Not EOF(1)
 Input #1, msg404
 Loop
 Close #1
 Exit Sub
new404:
 Open "C:\404.html" For Output As #1
 Print #1, "<h1><i>404 Page could not be found</i></h1>"
 Close #1
 
End Sub

Private Sub ListView1_Click()
Form2.Show
Form2.RichTextBox1.LoadFile Text1.Text & ListView1.SelectedItem.Text

End Sub

Private Sub Text1_Change()
On Error Resume Next
ListView1.ListItems.Clear


File1.Path = Text1.Text
filecount = File1.ListCount
For i = 1 To filecount - 1
  If UCase(Right(File1.List(i), 1)) = "L" Or UCase(Right(File1.List(i), 1)) = "M" Then ListView1.ListItems.Add i, "", File1.List(i), , 1
  If UCase(Right(File1.List(i), 1)) = "F" Or UCase(Right(File1.List(i), 1)) = "G" Then ListView1.ListItems.Add i, "", File1.List(i), , 2
  
  
 Next i
 HomePageFile = Text1.Text & "index.html"
End Sub

Private Sub Timer1_Timer()
'This is the one second timeout
'if the client doesn't request a new document in less than a second
'then it is probably done. If it is, dis-connect so somebody else can connect

Winsock1.Close
Winsock1.Listen

End Sub

Private Sub Winsock1_Close()
Winsock1.Close
Winsock1.Listen

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.Close
Winsock1.Accept requestID
List1.AddItem Winsock1.RemoteHostIP & " connected at " & Time
aryBlockedIPs = Split(BlockedIPs, " ", -1, vbBinaryCompare)
On Error GoTo AllsFine
'Get blocked ip list
Open "C:\blockedIPs.txt" For Input As #1
Do While Not EOF(1)
Input #1, BlockedIPs
Loop
Close #1


Exit Sub
AllsFine:
End Sub
Private Sub Blocked()
sResponse = "HTTP/1.1 200 OK" & vbNewLine & _
"Date: Sat, 02 Feb 2002 15:57:05 GMT" & vbNewLine & _
"Server: MWS/1.11" & vbNewLine & _
"Content-Type: text/html" & vbNewLine & _
"Content-Length: " & 73 & vbNewLine & _
"Cache-Control: private"
DoEvents
Winsock1.SendData sResponse & vbNewLine & vbNewLine & "<h1>You are not authorized to view this page</h1>Your IP has been blocked" & vbNewLine
List1.AddItem Winsock1.RemoteHostIP & " Blocked at " & Time
DoEvents
Winsock1.Close
Winsock1.Listen
DoEvents

End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
DoEvents
Winsock1.GetData Data, vbString, bytesTotal

'Check for blocked IP addresses
aryBlockedIPs = Split(BlockedIPs, " ", -1, vbBinaryCompare)
For Each ip In aryBlockedIPs
  If Winsock1.RemoteHostIP = ip Then
  Blocked
  Exit Sub
  End If
  
Next ip

Text2.Text = Data & Text2.Text


Timer1.Enabled = False
'The below chunk is checking to see if the user is connecting to the site for the first tiem;
'(like http://123.456.789.012/
DoEvents
If Left(Data, 14) = "GET / HTTP/1.1" Then   'First Connection
RichTextBox1.LoadFile HomePageFile
sResponse = "HTTP/1.1 200 OK" & vbNewLine & _
"Date: Sat, 02 Feb 2002 15:57:05 GMT" & vbNewLine & _
"Server: MWS/1.0" & vbNewLine & _
"Content-Type: text/html" & vbNewLine & _
"Content-Length: " & Len(RichTextBox1.Text) & vbNewLine & _
"Cache-Control: private"
List1.AddItem Winsock1.RemoteHostIP & " requested homepage"
Winsock1.SendData sResponse & vbNewLine & vbNewLine & RichTextBox1.Text & vbNewLine



End If

'This will load and send a picture (jpeg or gif) when the client's browser requests one
'If there is an error loading an image or html document, show your 404 error page
On Error GoTo show404
arydata = Split(Data, vbNewLine, -1, vbBinaryCompare)
For i = 1 To Len(arydata(0)) - 3
  If UCase(Mid(arydata(0), i, 3)) = "GIF" Or UCase(Mid(arydata(0), i, 3)) = "JPG" Then 'User is requesting a file
    
    arydata2 = Split(arydata(0), " ", -1, vbBinaryCompare)
    imagepath = Text1.Text & Right(arydata2(1), Len(arydata2(1)) - 1)
    RichTextBox1.LoadFile imagepath
    Winsock1.SendData RichTextBox1.Text
    List1.AddItem Winsock1.RemoteHostIP & " requested " & imagepath
  End If
Next i

'This will get HTML files when the client requests them.
For i = 1 To Len(arydata(0)) - 3
  If UCase(Mid(arydata(0), i, 4)) = "HTML" Or UCase(Mid(arydata(0), i, 3)) = "HTM" Then 'User is requesting a file
    
    arydata2 = Split(arydata(0), " ", -1, vbBinaryCompare)
    imagepath = Text1.Text & Right(arydata2(1), Len(arydata2(1)) - 1)
    Open imagepath For Input As #1
    Do While Not EOF(1)
    Input #1, test
    Loop
    Close #1
    RichTextBox1.LoadFile imagepath
    Winsock1.SendData sResponse & vbNewLine & vbNewLine & RichTextBox1.Text
     List1.AddItem Winsock1.RemoteHostIP & " requested " & imagepath
  End If
Next i


'If a user tries to DL an exe, zip, or mp3, do this...
For i = 1 To Len(arydata(0)) - 3
  If UCase(Mid(arydata(0), i, 3)) = "EXE" Or UCase(Mid(arydata(0), i, 3)) = "ZIP" Or UCase(Mid(arydata(0), i, 3)) = "MP3" Then 'User is requesting a file
    arydata2 = Split(arydata(0), " ", -1, vbBinaryCompare)
    imagepath = Text1.Text & Right(arydata2(1), Len(arydata2(1)) - 1)
    RichTextBox1.LoadFile imagepath
    Winsock1.SendData RichTextBox1.Text
    List1.AddItem Winsock1.RemoteHostIP & " requested " & imagepath
  End If
Next i
DoEvents


Timer1.Enabled = True
Exit Sub
show404:

sResponse = "HTTP/1.1 200 OK" & vbNewLine & _
"Date: Sat, 02 Feb 2002 15:57:05 GMT" & vbNewLine & _
"Server: MWS/1.11" & vbNewLine & _
"Content-Type: text/html" & vbNewLine & _
"Content-Length: " & Len(msg404) & vbNewLine & _
"Cache-Control: private"
Winsock1.SendData sResponse & vbNewLine & vbNewLine & msg404 & vbNewLine
List1.AddItem Winsock1.RemoteHostIP & " got 404 error at " & Time





End Sub

