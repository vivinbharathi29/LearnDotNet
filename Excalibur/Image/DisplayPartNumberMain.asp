<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
<script language="JavaScript" src="../../_ScriptLibrary/jsrsClient.js"></script>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>CD Part Number and Kit Number</TITLE>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

//-->
</SCRIPT>
</HEAD>

<BODY bgcolor=Ivory>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">

<font face=verdana size=2><b>
CD part Number and Kit Number
</b><br><br>
<%
	dim cn
	dim rs

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Application("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
		'Check for valid Root ID
		if request("RootID") = "" and request("VersionID") <> "" then
			rs.Open "spGetRootID " & clng(request("VersionID")),cn,adOpenForwardOnly
			if rs.EOF and rs.BOF then
				strRootID = rs("ID") & ""
			end if
			rs.Close
		else
			strRootID = request("RootID")
		end if
	
		rs.Open " spGetVersionProperties4Web " & clng(Request("VersionID")),cn,adOpenForwardOnly
		if not rs.EOF and not rs.BOF then
			strCDPart = rs("CDPartNumber") & ""
			strKitNumber = rs("CDKitNumber") & ""
			strMultiLanguage = rs("MultiLanguage") & ""			

			dim strPartRow
			dim strKitRow
			'Load Langs Supported By Version
			rs.Close

			rs.open "spGetSelectedLanguages "  & clng(request("VersionID")),cn,adOpenForwardOnly

			strPartRow=""
			strKitRow=""
			do while not rs.EOF
				if rs("ID") & "" <> "58" then
					strPartRow = strPartRow & "<TR ID=PNRow" & rs("Abbreviation") & "><TD nowrap>" & rs("Abbreviation") & " - " & rs("Name") & "</TD><TD><INPUT DISABLED style=""width=100%"" type=""text"" id=""txtPN" & rs("ID") & """ name=""txtPN" & rs("ID") & """ value=""" & rs("PartNumber") & """></TD></TR>"
					strKitRow = strKitRow & "<TR ID=KitRow" & rs("Abbreviation") & "><TD nowrap>" & rs("Abbreviation") & " - " & rs("Name") & "</TD><TD><INPUT DISABLED style=""width=100%"" type=""text"" id=""txtKit" & rs("ID") & """ name=""txtKit" & rs("ID") & """ value=""" & rs("CDKitNumber") & """></TD></TR>"	
				end if
				rs.MoveNext
			loop

			'Load Langs Supported By Root
			rs.Close
			rs.open "spGetLanguagesForRoot "  & clng(strRootID),cn,adOpenForwardOnly
			strAvailableLangs = ""
			do while not rs.EOF
				if rs("ID") & "" <> "58" then
					strPartRow = strPartRow & "<TR ID=PNRow" & left(rs("name") & "",2) & " style=""Display:none""><TD nowrap>" &  rs("Name") & "</TD><TD><INPUT DISABLED style=""width=100%"" type=""text"" id=""txtPN" & rs("ID") & """ name=""txtPN" & rs("ID") & """></TD></TR>"
					strKitRow = strKitRow & "<TR ID=KitRow" & left(rs("name") & "",2) & " style=""Display:none""><TD nowrap>" &  rs("Name") & "</TD><TD><INPUT DISABLED style=""width=100%"" type=""text"" id=""txtKit" & rs("ID") & """ name=""txtKit" & rs("ID") & """></TD></TR>"	
				end if
				rs.MoveNext
			loop		
			rs.Close
	end if
%>
	<table ID="tabGeneral" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
		<tr ID=CDPartNumberRow>
		<td nowrap valign=top>CD/DVD Part Number:</td>
		<td>
		<% if strMultiLanguage = "1" then%>
			<input DISABLED id="txtCDPart" name="txtCDPart" style="WIDTH: 250px; HEIGHT: 22px" size="2" maxlength="71" value="<%=strCDPart%>">
			<TABLE style="Display:none" ID=TablePart width=100%>
				<%=strPartRow%>
			</TABLE>
		<%else%>
			<input DISABLED id="txtCDPart" name="txtCDPart" style="Display:none;WIDTH: 100px; HEIGHT: 22px" size="2" maxlength="71" value="<%=strCDPart%>">
			<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 143px; BACKGROUND-COLOR: white" id=DIV1>
			<TABLE ID=TablePart width=100%>
			<THEAD><TR style="position:relative;top:expression(document.getElementById('DIV1').scrollTop-2);"><TD  bgcolor=lightsteelblue nowrap style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">Language</TD><TD  bgcolor=lightsteelblue style="width=100%; BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;PartNumber(s)&nbsp;</TD></TR></THEAD>
				<%=strPartRow%>
			</TABLE> 
			</div>
		<%end if%>					
		</td></tr>

		<tr ID=CDKitNumberRow>
		<td nowrap valign=top>Kit Number:</td>
		<td>
		<% if strMultiLanguage = "1" then%>
			<input DISABLED id="txtKitNumber" name="txtKitNumber" style="WIDTH: 250px; HEIGHT: 22px" size="2" maxlength="71" value="<%=strKitNumber%>">
			<TABLE style="Display:none" ID=TableKit width=100%>
				<%=strKitRow%>
			</TABLE>
		<%else%>
			<input DISABLED id="txtKitNumber" name="txtKitNumber" style="Display:none;WIDTH: 100px; HEIGHT: 22px" size="2" maxlength="71" value="<%=strKitNumber%>">
			<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 143px; BACKGROUND-COLOR: white" id=DIVKit>
			<TABLE ID=TableKit width=100%>
				<THEAD><TR style="position:relative;top:expression(document.getElementById('DIVKit').scrollTop-2);"><TD  bgcolor=lightsteelblue nowrap style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">Language</TD><TD  bgcolor=lightsteelblue style="width=100%; BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;KitNumber(s)&nbsp;</TD></TR></THEAD>
				<%=strKitRow%>
			</TABLE>
			</div>	
		<%end if%>					
		</td></tr>
	</table>

<%
	set rs=nothing
	cn.Close
	set cn=nothing
%>

</BODY>
</HTML>
