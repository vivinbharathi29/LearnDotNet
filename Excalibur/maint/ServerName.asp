<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<H2>SERVER_NAME</H2>
<P><%= Request.ServerVariables("SERVER_NAME")%></P>
<TABLE BORDER="1">
<TR><TD><B>Server Variable</B></TD><TD><B>Value</B></TD></TR>
<% For Each strKey In Request.ServerVariables %> 
<TR><TD> <%= strKey %> </TD><TD>  <%= Request.ServerVariables(strKey) %> </TD></TR>
<% Next %>
</TABLE>
</BODY>
</HTML>
