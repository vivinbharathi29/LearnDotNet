<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function optAll_onclick() {
	if (typeof(frmMain.chkDistribution.length) == "undefined")
		{
		frmMain.chkDistribution.checked=false;
		}
	else
		{
		for (i=0;i<frmMain.chkDistribution.length;i++)
			frmMain.chkDistribution(i).checked=false;
		}
	if (typeof(frmMain.chkTargetNotes.length) == "undefined")
		{
		frmMain.chkTargetNotes.checked=false;
		}
	else
		{
		for (i=0;i<frmMain.chkTargetNotes.length;i++)
			frmMain.chkTargetNotes(i).checked=false;
		}
	if (typeof(frmMain.chkImages.length) == "undefined")
		{
		frmMain.chkImages.checked=false;
		}
	else
		{
		for (i=0;i<frmMain.chkImages.length;i++)
			frmMain.chkImages(i).checked=false;
		}				
	CommentsRow.style.display="";
}

function optRemove_onclick(){
	optAll_onclick();
	optSelected_onclick();
	CommentsRow.style.display="none";
	frmMain.txtComments.value="";
}

function optSelected_onclick() {
	frmMain.chkTargetNotesAll.checked = false;
	frmMain.chkImagesAll.checked = false;
	frmMain.chkDistributionAll.checked = false;
	CommentsRow.style.display="";
}

function chkDistributionAll_onclick() {
	if (!frmMain.optAll.checked)
		{
		frmMain.optAll.checked = true;
		optAll_onclick();	
		}
}

function chkImagesAll_onclick() {
	if (!frmMain.optAll.checked)
		{
		frmMain.optAll.checked = true;
		optAll_onclick();	
		}
}

function chkTargetNotesAll_onclick() {
	if (!frmMain.optAll.checked)
		{
		frmMain.optAll.checked = true;
		optAll_onclick();	
		}
}

function chkDistribution_onclick() {
	if (!frmMain.optSelected.checked)
		{
		frmMain.optSelected.checked = true;
		optSelected_onclick();	
		}
}

function chkImages_onclick() {
	if (!frmMain.optSelected.checked)
		{
		frmMain.optSelected.checked = true;
		optSelected_onclick();	
		}
}

function chkTargetNotes_onclick() {
	if (!frmMain.optSelected.checked)
		{
		frmMain.optSelected.checked = true;
		optSelected_onclick();	
		}
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
	dim strVersionIDList
	dim strSQL
	dim VersionCount
	dim blnRootDistribution
	dim blnRootImages
	dim blnRootNotes
	dim strComments
	dim strCheckRootDistribution
	dim strCheckRootImages
	dim strCheckRootNotes
	dim blnHasRootExceptions
	dim blnHasVersionExceptions
	dim strCheckRootExceptions
	dim strCheckVersionExceptions
	
	strVersionIDList=""
	
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
		
		dim blnNeedsSupport
		rs.Open "spGetLeadProductRootExceptions " & clng(request("ProductID")) & "," & clng(request("RootID")) & "," & clng(request("ReleaseID")),cn,adOpenStatic
		if rs.EOF and rs.BOF then
			blnRootDistribution = false
			blnRootImages = false
			blnRootNotes = false
			strComments= ""
			blnNeedsSupport=true
		else
			blnNeedsSupport=false
			blnRootDistribution = rs("SyncDistribution")
			blnRootImages = rs("SyncImages")
			blnRootNotes = rs("SyncNotes")
			strComments= rs("SyncComments") & ""
		end if
		rs.Close
		blnHasRootExceptions = false
		if blnRootDistribution then
			strCheckRootDistribution = ""
		else
			strCheckRootDistribution = "checked"
			blnHasRootExceptions = true
		end if
		if blnRootImages then
			strCheckRootImages = ""
		else
			strCheckRootImages = "checked"
			blnHasRootExceptions = true
		end if
		if blnRootNotes then
			strCheckRootNotes = ""
		else
			strCheckRootNotes = "checked"
			blnHasRootExceptions = true
		end if
			
        if clng(request("ReleaseID")) = 0 then
		    strSQL = "SELECT v.ID, v.Version, v.Revision, v.Pass, pd.syncimages, pd.syncdistribution, pd.syncnotes, 0 as NeedsSupport " & _
				     "FROM DeliverableVersion v with (NOLOCK) inner join " & _
                     "product_deliverable pd with (NOLOCK) on v.id= pd.deliverableversionid " & _
				     "WHERE v.id in (" & scrubsql(mid(request("VersionIDList"),2)) & ") " & _
				     "and pd.productversionid = " & clng(request("ProductID")) & " "  & _
				     " UNION " & _
				     "SELECT v.ID, v.Version, v.Revision, v.Pass, pd.syncimages, pd.syncdistribution, pd.syncnotes, 1 as NeedsSupport " & _
				     "FROM DeliverableVersion v with (NOLOCK) inner join " & _
                     "product_deliverable pd with (NOLOCK) on v.id= pd.deliverableversionid inner join " & _
                     "productversion p with (NOLOCK) on p.referenceid = pd.productversionid " & _
				     "WHERE v.id in (" & scrubsql(mid(request("VersionIDList"),2)) & ") " & _
				     "and p.id = " & clng(request("ProductID")) & " "  & _
				     "and v.id not in ( " & _
					    "SELECT v.ID " & _
					    "FROM DeliverableVersion v with (NOLOCK) inner join " & _
                        "product_deliverable pd with (NOLOCK) on v.id= pd.deliverableversionid " & _
					    "WHERE v.id in (" & scrubsql(mid(request("VersionIDList"),2)) & ") " & _
					    "and pd.productversionid = " & clng(request("ProductID")) & ") " & _
				     "ORDER BY v.id desc;"
        else 
                strSQL = "SELECT v.ID, v.Version, v.Revision, v.Pass, pdr.syncimages, pdr.syncdistribution, pdr.syncnotes, 0 as NeedsSupport " & _
				     "FROM DeliverableVersion v with (NOLOCK) inner join " & _
                     "product_deliverable pd with (NOLOCK) on v.id= pd.deliverableversionid inner join " & _
                     "product_deliverable_Release pdr with (NOLOCK) on pdr.ProductDeliverableID =  pd.id " & _
				     "WHERE v.id in (" & scrubsql(mid(request("VersionIDList"),2)) & ") " & _
                     "and pdr.ReleaseID = " & clng(request("ReleaseID")) & " "  & _
				     "and pd.productversionid = " & clng(request("ProductID")) & " "  & _
				     " UNION " & _
				     "SELECT v.ID, v.Version, v.Revision, v.Pass, pdr.syncimages, pdr.syncdistribution, pdr.syncnotes, 1 as NeedsSupport " & _
				     "FROM DeliverableVersion v with (NOLOCK) inner join " & _
                     "product_deliverable pd with (NOLOCK) on v.id= pd.deliverableversionid inner join " & _
                     "product_deliverable_Release pdr with (NOLOCK) on pdr.ProductDeliverableID =  pd.id inner join " & _
                     "productversion_release lpvr with (NOLOCK) on lpvr.productversionid = pd.productversionid and pdr.ReleaseID = lpvr.ReleaseID inner join " & _
                     "productversion_release pvr with (NOLOCK) on lpvr.ID = pvr.LeadProductReleaseID " & _
				     "WHERE v.id in (" & scrubsql(mid(request("VersionIDList"),2)) & ") " & _
				     "and pvr.ProductVersionID = " & clng(request("ProductID")) & " "  & _
                     "and pvr.ReleaseID = " & clng(request("ReleaseID")) & " "  & _
				     "and v.id not in ( " & _
					    "SELECT v.ID " & _
					    "FROM DeliverableVersion v with (NOLOCK) inner join " & _
                        "product_deliverable pd with (NOLOCK) on v.id= pd.deliverableversionid inner join " & _
                        "product_deliverable_release pdr with (NOLOCK) on pdr.ProductDeliverableID = pd.ID " & _
					    "WHERE v.id in (" & scrubsql(mid(request("VersionIDList"),2)) & ") " & _
					    "and pd.productversionid = " & clng(request("ProductID")) & " " & _
                        "and pdr.ReleaseID = " & clng(request("ReleaseID")) & ") " & _
				     "ORDER BY v.id desc;"
        end if

		VersionCount = 0
		rs.Open strSQL,cn,adOpenStatic
		if rs.EOF and rs.BOF then
			strVersionList = ""
		else
			blnHasVersionExceptions = false
			do while not rs.EOF
				VersionCount = VersionCount + 1
				strVersionIDList = strVersionIDList & "," & trim(rs("ID"))
				strVersionList = strVersionList  & "<TR><TD>" & rs("Version")'& "<INPUT type=""checkbox"" id=chkVersions name=chkVersions LANGUAGE=javascript onclick=""return chkVersion_onclick()"" value=""" & rs("ID") & """>&nbsp;" & rs("Version")
				if trim(rs("Revision") & "") <> "" then
					strVersionList = strVersionList & "," & rs("Revision")
				end if
				if trim(rs("Pass") & "") <> "" then
					strVersionList = strVersionList & "," & rs("Pass")
				end if
				strVersionList = strVersionList & "</TD>"
				if rs("syncimages")=0 and not blnHasRootExceptions then
					SyncImagesChecked = "checked"
					blnHasVersionExceptions = true
				else
					SyncImagesChecked = ""
				end if
				if rs("syncdistribution")=0 and not blnHasRootExceptions then
					SyncDistributionChecked = "checked"
					blnHasVersionExceptions = true
				else
					SyncDistributionChecked = ""
				end if
				if rs("syncnotes")=0 and not blnHasRootExceptions then
					SyncNotesChecked = "checked"
					blnHasVersionExceptions = true
				else
					SyncNotesChecked = ""
				end if
				strVersionList = strVersionList & "<TD align=middle><INPUT " & SyncDistributionChecked & " type=""checkbox"" id=chkDistribution name=chkDistribution value=""" & rs("ID") & """ LANGUAGE=javascript onclick=""return chkDistribution_onclick()""></TD>"
				strVersionList = strVersionList & "<TD align=middle><INPUT " & SyncNotesChecked & " type=""checkbox"" id=chkTargetNotes name=chkTargetNotes value=""" & rs("ID") & """ LANGUAGE=javascript onclick=""return chkTargetNotes_onclick()""></TD>"
				strVersionList = strVersionList & "<TD align=middle><INPUT " & SyncImagesChecked & " type=""checkbox"" id=chkImages name=chkImages value=""" & rs("ID") & """ LANGUAGE=javascript onclick=""return chkImages_onclick()""></TD>"
				if rs("NeedsSupport") then
					strVersionList = strVersionList & "<TD><font color=red>This Root and Version will be added to " & strProductName & ".</font></TD>"
				else
					strVersionList = strVersionList & "<TD align=middle>none</TD>"
				end if
				strVersionList = strVersionList & "</TR>"
				rs.MoveNext
			loop
		end if
		rs.Close		

					
		if (not blnHasRootExceptions) and (not blnHasVersionExceptions) then
			strCheckRootExceptions = "checked"
			strCheckVersionExceptions = ""		
		elseif blnHasRootExceptions then
			strCheckRootExceptions = "checked"
			strCheckVersionExceptions = ""		
		else
			strCheckRootExceptions = ""
			strCheckVersionExceptions = "checked"			
		end if


		
		if strVersionIDList <> "" then
			strVersionIDList = mid(strVersionIDList,2)
		end if

%>


<h3>Edit Usage Exceptions</h3>
<P>Choose&nbsp;Fields&nbsp;to&nbsp;Ignore&nbsp;when&nbsp;Synchronizing&nbsp;<%=strProductName%>&nbsp;to&nbsp;Lead&nbsp;Product</P>
<form id=frmMain action="ExceptionsSave.asp" method=post>
<font size=2 face=verdana><b><%=strRootName%></b></font>
<table width="100%">
<TR>
	<TD colspan=2><INPUT <%=strCheckRootExceptions%> type ="radio" id=optAll name=optType value="1" LANGUAGE=javascript onclick="return optAll_onclick()">&nbsp;Apply&nbsp;Exceptions&nbsp;to&nbsp;All&nbsp;Versions&nbsp;(Current&nbsp;and&nbsp;Future&nbsp;Releases)</TD>
</TR>
<TR><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD width="100%">
	<%if blnNeedsSupport then
		Response.Write "<font color=red>&nbsp;This root will be added to " & strProductName & ".</font><BR>"
	  end if
	%>
	<INPUT <%=strCheckRootDistribution%> type="checkbox" id=chkDistributionAll name=chkDistributionAll LANGUAGE=javascript onclick="return chkDistributionAll_onclick()">&nbsp;Distribution<BR>
	<INPUT <%=strCheckRootNotes%> type="checkbox" id=chkTargetNotesAll name=chkTargetNotesAll LANGUAGE=javascript onclick="return chkTargetNotesAll_onclick()">&nbsp;Target Notes<BR>
	<INPUT <%=strCheckRootImages%> type="checkbox" id=chkImagesAll name=chkImagesAll LANGUAGE=javascript onclick="return chkImagesAll_onclick()">&nbsp;Image Summary<BR>
	</TD>
</TR>
<TR>
	<TD colspan=2><INPUT <%=strCheckVersionExceptions%> type="radio" id=optSelected name=optType value="2" LANGUAGE=javascript onclick="return optSelected_onclick()">&nbsp;Apply&nbsp;Exceptions&nbsp;to&nbsp;Selected&nbsp;Versions&nbsp;Only</TD>
</TR>
<TR><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD width="100%">
		<table border=1 cellpadding=2 cellspacing=1>
			<TR bgcolor=beige><TD><b>Version</b></TD><TD><b>Distribution</b></TD><TD><b>Target Notes</b></TD><TD><b>Image Summary</b></TD><TD><b>Warnings</b></TD></TR>
			<%=strVersionList%>
		</table>
	</TD>
</TR>
<%if blnHasRootExceptions or blnHasVersionExceptions then%>
	<TR>
<%else%>
	<TR style="display:none">
<%end if%>	
	<TD colspan=2><INPUT type="radio" id=optRemove name=optType value="3" LANGUAGE=javascript onclick="return optRemove_onclick()">&nbsp;Remove&nbsp;Exceptions&nbsp;from&nbsp;All&nbsp;Versions</TD>
</TR>

<TR ID=CommentsRow>
	<TD width="100%" colspan=2><BR>Comments:
	<TEXTAREA style="width:100%" rows=5 cols=80 id=txtComments name=txtComments><%=strComments%></TEXTAREA></TD>
</TR>
</table>
<INPUT type="hidden" id=txtProductID name=txtProductID value="<%=request("ProductID")%>">
<INPUT type="hidden" id=txtRootID name=txtRootID value="<%=request("RootID")%>">
<INPUT type="hidden" id=txtVersions name=txtVersions value="<%=strVersionIDList%>">
<INPUT type="hidden" id=txtReleaseID name=txtReleaseID value="<%=request("ReleaseID")%>">
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
