<%@ Language=VBScript %>
<!-- #include file = "includes/no-cache.asp" -->
<!-- #include file = "includes/noaccess.inc" -->
<HTML>
<HEAD>

<%if (Request.QueryString("Name") = "0" and Request.QueryString("NameType") = 0) then%>
    <TITLE>Add Service Tag</TITLE>
<%elseif (Request.QueryString("Name") = "0" and Request.QueryString("NameType") = 1) then%>
    <TITLE>Add BIOS Branding</TITLE>
<%elseif (Request.QueryString("Name") = "0" and Request.QueryString("NameType") = 2) then%>
    <TITLE>Add Logo Badge C Cover</TITLE>
<%elseif (Request.QueryString("Name") <> "0" and Request.QueryString("NameType") = 0) then%>
    <TITLE>Edit Service Tag</TITLE>
<%elseif (Request.QueryString("Name") <> "0" and Request.QueryString("NameType") = 1) then%>
    <TITLE>Edit BIOS Branding</TITLE>
<%elseif (Request.QueryString("Name") <> "0" and Request.QueryString("NameType") = 2) then%>
    <TITLE>Edit Logo Badge C Cover</TITLE>
<%end if%>

</HEAD>
	<FRAMESET ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="AddEditMarketingName.aspx?<%=Request.QueryString%>">
	</FRAMESET>
</HTML>