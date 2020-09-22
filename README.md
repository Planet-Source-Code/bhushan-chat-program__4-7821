<div align="center">

## Chat program


</div>

### Description

Chat program useing arrays.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bhushan\-](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bhushan.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__4-33.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bhushan-chat-program__4-7821/archive/master.zip)





### Source Code

```
1. //chatpage.asp
<HTML>
<HEAD>
<TITLE> Chat Page </TITLE>
</HEAD>
<FRAMESET ROWS="*,100">
<FRAME SRC="Display.asp">
<FRAME SRC="Message.asp">
</FRAMESET>
</HTML>
name this file as chatpage.asp. here we are dividing the page into 2 frames.
Next open another page and type the following code..
2. // Display.asp
<%
set MyServer=Request.ServerVariables("SERVER_NAME")
set MyPath=Request.ServerVariables("SCRIPT_NAME")
MySelf="HTTP://"&MyServer&MyPath
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="REFRESH" CONTENT="5;<%=MySelf%>">
<TITLE>Display Page</TITLE>
</HEAD>
<BODY>
<P ALIGN=RIGHT><%=NOW%></P>
<%
TempArray=Application("Talk")
FOR i=0 to Application("TPlace")-1
Response.Write("<P>"&Temparray(i))
NEXT
%>
</BODY>
</HTML>
name the page as display.asp. here the message typed by the user is captured in a Application level variable. Then the message is put in an array and is written on to the browser by reading from that array variable.
Next open another page and type the following code ..
3. // Message.asp
<%
IF not Request.Form("message")="" THEN
Application.LOCK
IF Application("TPlace")>4 THEN
Application("TPlace")=0
END IF
TempArray=Application("Talk")
TempArray(Application("TPlace"))=Request.Form("message")
Application("Talk")=TempArray
Application("TPlace")=Application("TPlace")+1
Application.Unlock
END IF
%>
<HTML>
<HEAD><TITLE> Message Page </TITLE></HEAD>
<BODY BGCOLOR="LIGHTBLUE">
<FORM METHOD="POST" ACTION="message.asp">
<p><INPUT TYPE=text size="50" name="message" >
<INPUT TYPE="SUBMIT" VALUE="SEND">
</p>
</FORM>
</BODY>
</HTML>
In this page we are creating a text field and a submit button. After typing a message if u click on submit button the message is captured in application variable which will be used by the display.asp
Next open notepad and type the following code and save it as global.asa
4. ///global.asa
<SCRIPT LANGUAGE=VBScript RUNAT=Server>
SUB Application_OnStart
Dim TempArray(5)
Application("Talk")=TempArray
Application("TPlace")=0
Application("ActiveUsers") = 0
END SUB
SUB Session_OnStart
' Change Session Timeout to 20 minutes (if you need to)
Session.Timeout = 20
' Set a Session Start Time
' This is only important to assure we start a session
Session ("Start") = Now
' Increase the active visitors count when we start the session
Application.Lock
Application ("ActiveUsers") = Application ("ActiveUsers") + 1
Application.UnLock
END SUB
SUB Session_OnEnd
' Decrease the active visitors count when the session ends.
Application.Lock
Application ("ActiveUsers") = Application ("ActiveUsers") - 1
Application.UnLock
END SUB
</SCRIPT>
```

