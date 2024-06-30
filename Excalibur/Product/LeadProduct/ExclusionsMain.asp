<%@ Language=VBScript %>
<!-- #include file = "../../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function optAll_onclick() {
	if (typeof(frmMain.chkVersions.length) == "undefined")
		{
		frmMain.chkVersions.checked=false;
		}
	else
		{
		for (i=0;i<frmMain.chkVersions.length;i++)
			frmMain.chkVersions(i).checked=false;
		}
}

function optSelected_onclick() {
	if (typeof(frmMain.chkVersions.length) == "undefined")
	{
		frmMain.chkVersions.checked=true;
	}
	else
	{
	    for (i = 0; i < frmMain.chkVersions.length; i++) {
	        frmMain.chkVersions(i).checked = true;
	    }
	}
}

function chkVersion_onclick(){
	frmMain.optSelected.checked = true;
}

//-->
</SCRIPT>
</HEAD>
<style>
h3{
	FONT-FAMILY: Verdana;
	FONT-SIZE: small;
}
body{
	FONT-FAMILY: Verdana;
	FONT-SIZE: x-small;
}
td{
	FONT-FAMILY: Verdana;
	FONT-SIZE: x-small;
}
</style>
<BODY bgColor=ivory>

<%
	dim cn,rs
	dim strRootName
	dim strProductName
	dim strVersionList
	dim strSQL
	dim VersionCount
	
	if request("RootID") = "" or request("ProductID") = "" or request("VersionIDList") = "" then
		Response.Write "Not enough information supplied to display the requested page."
	else

		set cn = server.CreateObject("ADODB.Connection")
		set rs = server.CreateObject("ADODB.recordset")
	
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
		cn.Open

		rs.Open "spGetDeliverableRootName " & clng(request("RootID")),cn,adOpenStatic
		if rs.EOF and rs.BOF then
			strRootName = ""
		else
			strRootName = rs("name") & ""
		end if
		rs.Close

		rs.Open "spGetProductVersionName " & clng(request("ProductID")) & "," & clng(request("ReleaseID")),cn,adOpenStatic
		if rs.EOF and rs.BOF then
			strProductName = ""
		else
			strProductName = rs("name") & ""
		end if
		rs.Close

		strSQL = "SELECT ID, Version, Revision, Pass " & _
				 "FROM DeliverableVersion with (NOLOCK) " & _
				 "WHERE ID in (" & scrubsql(mid(request("VersionIDList"),2)) & ") " & _
				 "ORDER BY id desc;"

		VersionCount = 0
		rs.Open strSQL,cn,adOpenStatic
		if rs.EOF and rs.BOF then
			strVersionList = ""
		else
			do while not rs.EOF
				VersionCount = VersionCount + 1
				strVersionList = strVersionList & "<INPUT type=""checkbox"" id=chkVersions name=chkVersions LANGUAGE=javascript onclick=""return chkVersion_onclick()"" value=""" & rs("ID") & """>&nbsp;" & rs("Version")
				if trim(rs("Revision") & "") <> "" then
					strVersionList = strVersionList & "," & rs("Revision")
				end if
				if trim(rs("Pass") & "") <> "" then
					strVersionList = strVersionList & "," & rs("Pass")
				end if
				strVersionList = strVersionList & "<BR>"
				rs.MoveNext
			loop
		end if
		rs.Close		
		
		

%>

<h3>Exclude Versions from Lead Product Synchronization</h3>

Exclude <%=strRootName%> synchronization on <%=strProductName%><BR><BR>
<form id=frmMain action=ExclusionsSave.asp method=post>
<table width="100%">
<TR>
	<TD colspan=2><INPUT type="radio" id=optAll name=optType value=1 CHECKED 
      LANGUAGE=javascript onclick="return optAll_onclick()">&nbsp;Exclude All Versions (Current and Future Releases)</TD>
</TR>
<TR>
	<TD  colspan=2><INPUT type="radio" id=optSelected name=optType value=2 LANGUAGE=javascript onclick="return optSelected_onclick()">&nbsp;Exclude Selected
	<%if VersionCount> 1 then%>
	 Versions Only
	<%else%>
	 Version Only
	<%end if%>
	</TD>
</TR>
<TR><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD width="100%">
	<%=strVersionList%>
	</TD>
</TR>
<TR>
	<TD width="100%" colspan=2><BR>Comments:
	<TEXTAREA style="width:100%" rows=5 cols=80 id=txtComments name=txtComments></TEXTAREA></TD>
</TR>
</table>
<BR><BR>

<INPUT style="Display:none" type="text" id=txtProductID name=txtProductID value="<%=request("ProductID")%>">
<INPUT style="Display:none" type="text" id=txtRootID name=txtRootID value="<%=request("RootID")%>">
<INPUT style="Display:none" type="text" id=txtReleaseID name=txtReleaseID value="<%=request("ReleaseID")%>">
</form>

<%
		set rs = nothing
		cn.Close
		set cn = nothing

	end if
%>

</BODY>
</HTML>

<%
	function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		dim i
		
		strWords=replace(strWords,"'","''")
		
		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "=", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 

%>
