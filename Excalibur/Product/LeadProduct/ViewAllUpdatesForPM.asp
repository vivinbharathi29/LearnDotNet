<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

	<table cellspacing="0" border="0" width="100%">
	<tr>
		<td colspan="5">
		<p>
		<font size="2" face="verdana"><strong><u>Lead&nbsp;Product&nbsp;-&nbsp;Synchronization&nbsp;Issues</u></strong></font></p></td>
		<td><font size="2"></font></td></tr>
  <%
		set cn = server.CreateObject("ADODB.Connection")
		set rs = server.CreateObject("ADODB.recordset")
	
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
		cn.Open
	strLeadSyncVersionIDList = ""
	if Request("PreinstallPM")="1" then
		rs.Open "spListLeadProductDiscrepencies " & Request("UserID") & ",1" ,cn,adOpenStatic,adLockReadOnly ' & clng(CurrentUserID)
	else
		rs.Open "spListLeadProductDiscrepencies " & Request("UserID") & ",2" ,cn,adOpenStatic,adLockReadOnly ' & clng(CurrentUserID)
	end if
	if rs.EOF and rs.BOF then
		Response.Write "<TR><TD><font size=1 face=verdana><b>none</b></font></TD></TR>"
	else
		Response.Write "<TR><TD colspan=5><BR><font size=1 color=red face=verdana>Note: This is a read-only view of all lead product sync alerts assigned to this PM.<BR><BR></font></TD></TR>"
  %>
  <tr bgcolor="beige">
	<td width="130"><font face="verdana" size="1"><b>Product&nbsp;&nbsp;&nbsp;</b></font></td>
	<td width="130"><font face="verdana" size="1"><b>Lead&nbsp;Product&nbsp;&nbsp;&nbsp;</b></font></td>
	<td width="35%"><font face="verdana" size="1"><b>Deliverable</b></font></td>
	<td width="65%"><font face="verdana" size="1"><b>Actions</b></font></td>
  </tr>
  <%
		

	i=0
	LastRoot = 0
	LastFollower=0
	do while not rs.EOF
			strVersion = "["& rs("Version")
			if rs("Revision") <> "" then
				strVersion = strVersion & "," & rs("Revision")
			end if
			if rs("Pass") <> "" then
				strVersion = strVersion & "," & rs("Pass")
			end if
			strVersion = strVersion & "]"
			strLeadproductHeaderRowBGColor = "gainsboro"
			strLeadproductDataRowBGColor = "white"
			
			
			strLeadDistribution = ""
			if rs("LeadPreinstall") then
				strLeadDistribution = strLeadDistribution & ", Preinstall"
				if trim(rs("LeadPreinstallBrand")&"") <> "" then
					strLeadDistribution = strLeadDistribution & " (" & trim(rs("LeadPreinstallBrand")) & ")"
				end if
			end if
			if rs("LeadPreload") then
				strLeadDistribution = strLeadDistribution & ", Preload"
				if trim(rs("LeadPreloadBrand")&"") <> "" then
					strLeadDistribution = strLeadDistribution & " (" & trim(rs("LeadPreloadBrand")) & ")"
				end if
			end if
			if rs("LeadWeb") then
				strLeadDistribution = strLeadDistribution & ", Web"
			end if
			if rs("LeadDropInBox") then
				strLeadDistribution = strLeadDistribution & ", DIB"
			end if
			if rs("LeadSelectiveRestore") then
				strLeadDistribution = strLeadDistribution & ", SelectiveRestore"
			end if
			if rs("LeadARCD") then
				strLeadDistribution = strLeadDistribution & ", DRCD"
			end if
			if rs("LeadDRDVD") then
				strLeadDistribution = strLeadDistribution & ", DRDVD"
			end if
			if rs("LeadRACD_EMEA") then
				strLeadDistribution = strLeadDistribution & ", RACD_EMEA"
			end if
			if rs("LeadRACD_APD") then
				strLeadDistribution = strLeadDistribution & ", RACD_APD"
			end if
			if rs("LeadRACD_Americas") then
				strLeadDistribution = strLeadDistribution & ", RACD_Americas"
			end if
			if rs("LeadDocCD") then
				strLeadDistribution = strLeadDistribution & ", DocCD"
			end if
			if rs("LeadOSCD") then
				strLeadDistribution = strLeadDistribution & ", OSCD"
			end if
			if trim(rs("LeadPatch")&"") <> "0" then
				strLeadDistribution = strLeadDistribution & ", Patch"
			end if
			if strLeadDistribution <> "" then
				strLeadDistribution = mid(strLeadDistribution,3)
			end if


			strFollowDistribution = ""
			if rs("FollowPreinstall") then
				strFollowDistribution = strFollowDistribution & ", Preinstall"
				if trim(rs("FollowPreinstallBrand")&"") <> "" then
					strFollowDistribution = strFollowDistribution & " (" & trim(rs("FollowPreinstallBrand")) & ")"
				end if
			end if
			if rs("FollowPreload") then
				strFollowDistribution = strFollowDistribution & ", Preload"
				if trim(rs("FollowPreloadBrand")&"") <> "" then
					strFollowDistribution = strFollowDistribution & " (" & trim(rs("FollowPreloadBrand")) & ")"
				end if
			end if
			if rs("FollowWeb") then
				strFollowDistribution = strFollowDistribution & ", Web"
			end if
			if rs("FollowDropInBox") then
				strFollowDistribution = strFollowDistribution & ", DIB"
			end if
			if rs("FollowSelectiveRestore") then
				strFollowDistribution = strFollowDistribution & ", SelectiveRestore"
			end if
			if rs("FollowARCD") then
				strFollowDistribution = strFollowDistribution & ", DRCD"
			end if
			if rs("FollowDRDVD") then
				strFollowDistribution = strFollowDistribution & ", DRDVD"
			end if
			if rs("FollowRACD_EMEA") then
				strFollowDistribution = strFollowDistribution & ", RACD_EMEA"
			end if
			if rs("FollowRACD_APD") then
				strFollowDistribution = strFollowDistribution & ", RACD_APD"
			end if
			if rs("FollowRACD_Americas") then
				strFollowDistribution = strFollowDistribution & ", RACD_Americas"
			end if
			if rs("FollowDocCD") then
				strFollowDistribution = strFollowDistribution & ", DocCD"
			end if
			if rs("FollowOSCD") then
				strFollowDistribution = strFollowDistribution & ", OSCD"
			end if
			if trim(rs("FollowPatch")&"") <> "0" then
				strFollowDistribution = strFollowDistribution & ", Patch"
			end if
			if strFollowDistribution <> "" then
				strFollowDistribution = mid(strFollowDistribution,3)
			end if

			if LastRoot <> rs("RootID") then			
				if lastroot <> 0 then
					i=i+1
					response.write "<INPUT type=""hidden"" id=txtLeadSyncVersionList" & trim(LastRoot) & "_" & trim(LastFollower) & " name=txtLeadSyncVersionList" & trim(LastRoot) & "_" & trim(LastFollower) & " value=""" & 	strLeadSyncVersionIDList &  """>"
					strLeadSyncVersionIDList = ""
					Response.Write "</TD></TR>"
					if i = 51 then
						Response.Write "<TR><TD colspan=4 bgcolor=Wheat><b><font size=2 face=verdana>The following alerts are not displayed on the today page.</b></font></TD></TR>"
					end if
				end if
				LastRoot = rs("RootID")
				LastFollower = rs("FollowerID")
				%>
				</tr>
				<!--class="VersionID=<%=rs("VersionID")%>&amp;ProductID=<%=rs("FollowerID")%>"-->
				<tr bgcolor="ivory"  id="LeadSync<%=rs("FollowerID")& ":" & rs("RootID")%>">
					<td style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top" nowrap><font face="verdana" class="text" size="1"><%=rs("Product")%>&nbsp;&nbsp;</font></td>
					<td nowrap style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top"><font face="verdana" class="text" size="1"><%=rs("leadProduct")%>&nbsp;</font></td>
					<td style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top"><font face="verdana" class="text" size="1"><%=rs("DeliverableName") %>&nbsp;</font></td>
					<TD style="BORDER-TOP: <%=strRowBorderColor%> thin solid" class="cell" valign="top"><font face="verdana" class="text" size="1">
	<!--			
	<TR bgcolor=ivory><TD>&nbsp;</TD><TD bgcolor=<%=strLeadproductHeaderRowBGColor%>><b>Action</b></TD><TD bgcolor=<%=strLeadproductHeaderRowBGColor%>><b>Version</b></TD><TD bgcolor=<%=strLeadproductHeaderRowBGColor%>><b>Lead Distribution</b></TD><TD bgcolor=<%=strLeadproductHeaderRowBGColor%>><b>Lead Target Notes</b></TD><TD bgcolor=<%=strLeadproductHeaderRowBGColor%>><b>Lead Image Summary</b></TD></TR>
	-->
			<%
			end if
%>
		<!--		<TR bgcolor=ivory><TD>&nbsp;</TD><TD style="BORDER-LEFT: gainsboro thin solid" nowrap bgcolor=<%=strLeadproductDataRowBGColor%>><%=rs("ActionPlan")%></TD><TD nowrap bgcolor=<%=strLeadproductDataRowBGColor%>><%=strVersion%></TD><TD bgcolor=<%=strLeadproductDataRowBGColor%>><%=strDistribution%></TD><TD bgcolor=<%=strLeadproductDataRowBGColor%>><%=rs("LeadtargetNotes") & ""%></TD><TD bgcolor=<%=strLeadproductDataRowBGColor%>><%=rs("LeadImageSummary") & ""%></TD></TR>
				<TR bgcolor=ivory height=5><TD></TD><TD></TD><TD></TD><TD></TD><TD></TD><TD></TD></tr>
-->
<%
			if isnull(rs("LeadImageSummary")) then
				strLeadImageSummary = "All"
			elseif trim(rs("LeadImageSummary") & "") = "" then
				strLeadImageSummary = "All"
			else
				strLeadImageSummary = rs("LeadImageSummary")
			end if 
			
			if isnull(rs("Supported")) then
				strFollowImageSummary =  ""
			elseif isnull(rs("FollowImageSummary")) then
				strFollowImageSummary = "All"
			elseif trim(rs("FollowImageSummary") & "") = "" then
				strFollowImageSummary = "All"
			else
				strFollowImageSummary =  rs("FollowImageSummary") & ""
			end if
			
			strLeadSyncVersionIDList = strLeadSyncVersionIDList & "," & rs("VersionID")

			if (rs("LeadTargeted") and not rs("FollowTargeted")) or ((rs("LeadTargeted") and isnull(rs("FollowTargeted"))))then 
				Response.Write "-Target " & strversion &  "<BR>"
			end if
			if rs("Planorder") =2 then 
				Response.Write "-Remove Target " & strversion &  "<BR>"
			end if
			if strFollowDistribution <> strleadDistribution and (rs("SyncDistribution") or isnull(rs("Supported"))  ) then
				Response.Write "-Update Distribution "  & "on " & strversion & " to '" & strLeadDistribution  & "'<BR>"
			end if
			if rs("LeadtargetNotes")&""  <> rs("FollowtargetNotes")&"" and (rs("Syncnotes") or isnull(rs("Supported")) )then
				Response.Write "-Update TargetNotes "  & "on " & strversion & " to '" & rs("LeadtargetNotes")  & "'<BR>"
			end if
			if strLeadImageSummary  <> strFollowImageSummary and ( rs("SyncImages") or isnull(rs("Supported")) ) then
				Response.Write "-Update Image Summary "  & "on " & strversion & " to '" & strLeadImageSummary  & "'<BR>"
			end if
		rs.MoveNext
	loop
	i=i+1
	
	end if
	rs.Close
	set rs = nothing
	cn.Close
	set cn=nothing
	%><input type="hidden" id="txtLeadSyncCount" name="txtLeadSyncCount" value="<%=i%>">
	</table>


</BODY>
</HTML>
