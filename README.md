<div align="center">

## Amazing Winsock Chat Tutorial


</div>

### Description

This tutorial will walk you through the process of making a decent, and fully functional winsock chat. Extremely easy to comprehend, and made for newbies.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ Yariv Sarafraz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/yariv-sarafraz.md)
**Level**          |Beginner
**User Rating**    |5.0 (100 globes from 20 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/yariv-sarafraz-amazing-winsock-chat-tutorial__1-38262/archive/master.zip)





### Source Code

<body>
<p><font color="#FFFFFF">.</font></p>
<p><font face="Arial" size="2"><b>Winsock Chat Tutorial </b><font color="#808080">[Tutorial
by Yariv Sarafraz] </font><b><BR></b>First off, I've
seen many winsock chat tutorials and examples on psc, and most of them are
vague. This tutorial will walk you through the process of making a decent,
functional winsock chat. So let's get started ...</font></p>
<p align="left"><font face="Arial" size="2"> We'll start with the client code.</font></p>
<p><font face="Arial" size="2"><u>Client Code<BR></u>Requirements: Winsock
Control, 3 textboxes, 1 rich textbox, 1 label, and 2 command buttons. Let's get started ...<BR><BR>The
rich textbox will be the
chat room itself, so let's name it 'txtchat'.  We'll name the first textbox
'txtsend',
the second we'll name 'txtport', and the third we'll name 'txtip'. Name the label
'lblstatus'. Add the 2 command buttons, as well as, the winsock control to the form.</font></p>
<p><font face="Arial" size="2"><font color="#808080">Private Sub</font><font color="#FF0000"> Form_Load</font><font color="#808080">()</font><BR><font color="#000000">txtport.Text
= "616" </font><font color="#008080">'Assigning a port</font></font>
<BR><font size="2" face="Arial">txtip.Text = "127.0.0.1"</font><font color="#008080"><BR></font><font face="Arial" size="2"><font color="#808080">End
Sub</font></font></p>
<p><font size="2" face="Arial"><font color="#808080">Private Sub </font><font color="#FF0000">Winsock1_Connect</font><font color="#808080">()</font><br>
lblstatus.Caption  = "Connected to Server!"  <font color="#008080">'Reporting
status</font><br>
txtchat.Text = txtchat.Text & "*** Connection Complete Achieved ... ***" 
<font color="#008080">'Reporting status to chat</font><br>
<font color="#808080">End Sub</font></font>
</p>
<p><font face="Arial" size="2"><font color="#808080">Private Sub </font><font color="#FF0000">Winsock1_ConnectionRequest</font><font color="#808080">(ByVal requestID As Long)<br>
</font><font color="#000000">If Winsock1.State <> sckClosed Then Winsock1.Close 
</font><font color="#008080">'If the winsock control is in use, close it</font><font color="#000000"><br>
Winsock1.Accept requestID  </font><font color="#008080">'Allow
connection</font><font color="#808080"><br>
End Sub</font></font></p>
<p><font face="Arial" size="2"><font color="#808080">Private Sub </font><font color="#FF0000">Winsock1_Error</font><font color="#808080">(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)</font><br>
lblstatus.caption = "Error Occurred."  <font color="#000000">
</font><font color="#008080">'If an error occurs, report it to the status bar<BR></font>Command2.Enabled =
True<br>
<font color="#808080">End Sub</font></font></p>
<p><font face="Arial" size="2"><font color="#808080">Private Sub </font><font color="#FF0000">Winsock1_DataArrival</font><font color="#808080">(ByVal bytesTotal As Long)</font><br>
Dim Data As String<br>
Winsock1.GetData Data  <font color="#000000"> </font><font color="#008080">'Get
whatever data the server has sent to you (the client)</font><br>
txtchat.Text = txtchat.Text & vbNewLine & "Server: " & vbTab & Data 
<font color="#008080">'Whatever
data we receive, place into the chat textbox, txtchat.Text.</font><br>
<font color="#808080">End Sub</font></font></p>
<p><font face="Arial" size="2"><font color="#808080">Private Sub </font><font color="#FF0000">txtchat_Change</font><font color="#808080">()</font><br>
txtchat.SelStart = Len(txtchat.Text)  <font color="#008080">'This
one line of code will responds to the fact text is being sent to txtchat.Text,
and as text enters the textbox, it automatically scrolls down to the last line.
If you don't know what I'm talking about, just try the code and you'll
understand.</font><br>
<font color="#808080">End Sub</font></font></p>
<p><font face="Arial" size="2"><font color="#800080">'Make command button 1's
caption 'Send Text'</font><font color="#808080"><BR>Private Sub</font> <font color="#FF0000">Command1_Click</font><font color="#808080">()</font><br>
Data = txtsend.Text<br>
If
lblstatus.Caption  = "Connected to Server!"  Then  <font color="#008080">'If
the status bar states that you're connected, then ...</font><br>
txtchat.Text = txtchat.Text & vbNewLine & "Client: " & vbTab &
txtsend.Text  <font color="#008080">'Add the text you typed to the chatbox</font><br>
Winsock1.SendData (Data) <font color="#008080">'And send the text you typed to
the server</font><br>
txtsend.Text = "" <font color="#008080">'Clear txtsend.Text</font><br>
Else <font color="#008080">'But! If the status label says anything but
'Connected to Server!', do the following</font><br>
txtchat.Text = txtchat.Text & vbNewLine & "* Connection Lost/Not
Found." <font color="#008080">'Send to the chat window an error message</font><br>
End If<Br><font color="#808080">End Sub</font></font></p>
<p><font face="Arial" size="2"><font color="#800080">'Make command button 2's
caption 'Connect to Server'</font><font color="#808080"><BR>Private Sub</font> <font color="#FF0000">Command2_Click</font><font color="#808080">()<BR></font><font color="#000000">If
txtip.Text = "" then </font> <font color="#008080">'If the IP
text box (txtip.Text) is empty, then do the following ...</font><font color="#000000"><BR>Msgbox
"Please enter an IP number." </font> <font color="#008080">'Send
a message box to report the problem</font><font color="#000000"><BR>Exit Sub<BR>Else  </font>
<font color="#008080">'But if the IP text box (txtip.Text) is anything but
empty, do the following ...</font><font color="#808080"><BR></font>Command2.Enabled
= False<br>
 Winsock1.Close  <font color="#008080">'Close any current connection<BR></font><font color="#000000">Winsock1.Connect
txtip.text, txtport.text  </font><font color="#008080">'Connect to the given IP and port.</font><br>
lblstatus.Caption  = "Connecting ..."  <font color="#008080">'Report
status to status label</font><Br><font color="#808080">End Sub</font></font></p>
<p><font face="Arial" size="2"><font color="#808080">Private Sub </font><font color="#FF0000">Form_Unload</font><font color="#808080">(Cancel As Integer)<br>
</font>On Error Resume Next  <font color="#008080">'If
an error were to occur, ignore it and keep working</font><br>
Winsock1.SendData "Exit" <font color="#008080">'See
the Winsock1_DataArrival for the <i>server </i>to understand what this is doing
here</font><br>
Do Events <br>
Winsock1.Close  <font color="#008080">'Close
the connection</font><br>
End  <font color="#008080">'Exit the program</font><br>
<font color="#808080">End Sub</font></font></p>
<p><font color="#0000FF" size="2" face="Arial">------------------------------:
Client code ends here! We now begin the <i>Server </i>code.</font></p>
<p><font face="Arial" size="2"><u>Server Code<BR></u>Requirements: Winsock
Control, 1 textbox, 1 rich textbox, 1 label, and 2 command buttons. Let's get started ...<BR><BR>The
rich textbox will be the
chat room itself, so let's name it 'txtchat'.  We'll name the textbox
'txtsend'. Name the label
'lblstatus'. Add the 2 command buttons, as well as, the winsock control to the form.</font></p>
<p><font face="Arial" size="2"><font color="#800080">'Make command button 1's
caption 'Listen for Connection'</font><font color="#808080"><BR>Private Sub</font> <font color="#FF0000">Command1_Click</font><font color="#808080">()</font><br>
Winsock1.Close <font color="#000000"> </font><font color="#008080">'Close
any current winsock connection</font><br>
Winsock1.LocalPort = 616 <font color="#000000"> </font><font color="#008080">'Assigning
a port</font><br>
Winsock1.Listen<font color="#000000"> </font><font color="#008080">'Listen for
connection</font><BR>
lblstatus.Caption = "Listening for connection ..." <font color="#000000">
</font><font color="#008080">'Report status to status label</font>
<br>
txtchat.Text = txtchat.Text & "*** Not Connected. Please stand by ... ***"  <font color="#000000">
</font><font color="#008080">'Report status to chat</font><br>
End Sub</font></p>
<p><font face="Arial" size="2"><font color="#800080">'Make command button 2's
caption 'Send Text'</font><font color="#808080"><BR>Private Sub</font> <font color="#FF0000">Command2_Click</font><font color="#808080">()</font><br>
Data = txtsend.Text <font color="#008080"> 'Define 'Data'</font><br>
If
lblstatus.Caption  = "Connected."  Then  <font color="#008080">'If
the status bar states that you're connected, then ...</font><br>
txtchat.Text = txtchat.Text & vbNewLine & "Server: " & vbTab &
txtsend.Text  <font color="#008080">'Add the text you typed to the chatbox</font><br>
Winsock1.SendData (Data) <font color="#008080">'And send the text you typed to
the client</font><br>
txtsend.Text = "" <font color="#008080">'Clear txtsend.Text</font><br>
Else <font color="#008080">'But! If the status label says anything but
'Connected.', do the following</font><br>
txtchat.Text = txtchat.Text & vbNewLine & "* Connection Lost/Not
Found." <font color="#008080">'Send an error message to the chat window</font><br>
End If<br>
End Sub</font></p>
<p><font face="Arial" size="2"><font color="#808080">Private Sub </font><font color="#FF0000">Winsock1_ConnectionRequest</font><font color="#808080">(ByVal requestID As Long)<br>
</font><font color="#000000">If Winsock1.State <> sckClosed Then Winsock1.Close 
</font><font color="#008080">'If the winsock control is in use, close it</font><font color="#000000"><br>
Winsock1.Accept requestID  </font><font color="#008080">'Allow
connection<BR></font>lblstatus.Caption = "Connected."<font color="#808080"><br>
End Sub</font></font></p>
<p><font face="Arial" size="2"><font color="#808080">Private Sub </font><font color="#FF0000">Winsock1_Error</font><font color="#808080">(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)</font><br>
lblstatus.caption = "Error Occurred."  <font color="#000000">
</font><font color="#008080">'If an error occurs, report it to the status bar</font><br>
<font color="#808080">End Sub</font></font></p>
<p><font face="Arial" size="2"><font color="#808080">Private Sub </font><font color="#FF0000">Winsock1_DataArrival</font><font color="#808080">(ByVal bytesTotal As Long)</font><br>
Dim Data As String<br>
Winsock1.GetData Data  <font color="#008080">'Get
the data sent by the client</font><br>
If Data = "Exit" Then <font color="#008080">'If
the data string is 'Exit' then do the following ...</font><br>
Winsock1.Close <font color="#008080">'Close the winsock connection</font><br>
lblstatus.Caption = "Connection Forcefully Ended." <font color="#008080">'The
status bar will display the fact that the server is disconnected</font><br>
Else <font color="#008080">'But, If the data string is anything but
'Exit', then ...</font><br>
txtchat.Text = txtchat.Text & vbNewLine & "Client: " & vbTab & Data 
<font color="#008080">'Add data to txtchat.Text (the chatroom)</font><br>
End If<br>
<font color="#808080">End Sub</font></font></p>
<p><font face="Arial" size="2"><font color="#808080">Private Sub </font><font color="#FF0000">txtchat_Change</font><font color="#808080">()</font><br>
txtchat.SelStart = Len(txtchat.Text)  <font color="#008080">'This
one line of code will responds to the fact text is being sent to txtchat.Text,
and as text enters the textbox, it automatically scrolls down to the last line.
If you don't know what I'm talking about, just try the code and you'll
understand.</font><br>
<font color="#808080">End Sub</font></font></p>
<p align="center"><BR><font face="Arial" size="2"><BR><font color="#0000FF">-----------------------------------------------------------<BR><b>That's it! You're
done. It's really not that complicated.</b></font></font></p>
<p align="center"><font color="#800080" size="2" face="Arial">I am currently working on a RAT (Remote Admin Tool), and if you have any
knowledge concerning the programming of the 'Edit Server' part of a RAT, please
contact me, at:  <a href="mailto:wise@iamwasted.com">wise@iamwasted.com</a>,
or ICQ# 164953049</font></p>
<p align="center"><font color="#808080" size="2" face="Arial">[  <i>Enjoy the tutorial
-  Yariv Sarafraz  </i>]</font></p>
</body>

