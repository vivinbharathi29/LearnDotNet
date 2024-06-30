<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value == "1")
			{
			window.returnValue = 1;
			window.parent.close();
			}
		else
			document.write ("<BR><font size=2 face=verdana>Unable to update roadmap item.</font>");
		}
	else
		document.write ("<BR><font size=2 face=verdana>Unable to update roadmap item.</font>");
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">


<%
	dim cn
	dim rs
	dim cm
	dim p
	dim i
	dim strNewPriority
	dim strNewID
	dim strSuccess

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	'Save Roadmap Item
	set cm = server.CreateObject("ADODB.Command")
		
	cm.ActiveConnection = cn
	cm.CommandType = &H0004
	if trim(request("txtDisplayedID"))="" then
		cm.CommandText = "spAddActionRoadmapItem"
	else
		strNewID = request("txtDisplayedID")
		
		cm.CommandText = "spUpdateActionRoadmapItem"
		
		Set p = cm.CreateParameter("@ID",adInteger, &H0001)
		p.Value = clng(request("txtDisplayedID"))
		cm.Parameters.Append p

	end if

	Set p = cm.CreateParameter("@ProductVersionID",adInteger, &H0001)
	p.Value = clng(request("cboProject"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@OwnerID",adInteger, &H0001)
	p.Value = clng(request("cboOwner"))
	cm.Parameters.Append p
			
	Set p = cm.CreateParameter("@ActionStatusID",adInteger, &H0001)
	p.Value = clng(request("cboStatus"))
	cm.Parameters.Append p
					
	Set p = cm.CreateParameter("@Summary",adVarChar, &H0001,256)
	p.Value = left(request("txtSummary"),256)
	cm.Parameters.Append p
					
	if trim(request("txtDisplayedID"))<>"" then
		Set p = cm.CreateParameter("@OriginalTimeframe",adVarChar, &H0001,30)
		if trim(request("txtOriginalTimeframe")) = "" and trim(request("txtTimeframe")) <> "" then
			p.Value = left(request("txtTimeframe"),30)
		else
			p.value = left(request("txtOriginalTimeframe"),30)
		end if
		cm.Parameters.Append p
	end if
						
	Set p = cm.CreateParameter("@Timeframe",adVarChar, &H0001,30)
	p.Value = left(request("txtTimeframe"),30)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@TimeframeNotes",adVarChar, &H0001,500)
	p.Value = left(request("txtTimeframeNotes"),500)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@DisplayOrder",adInteger, &H0001)
	p.Value = 0
	cm.Parameters.Append p

	if trim(request("txtDisplayedID"))<>"" then
		Set p = cm.CreateParameter("@StatusReport", 16, &H0001)
		if request("chkReport") = "on" then
			p.value = 1
		else
			p.value = 0
		end if
		cm.Parameters.Append p
	end if
	
	Set p = cm.CreateParameter("@Notes",adVarChar, &H0001,80)
	p.Value = left(request("txtNotes"),80)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Details",adLongVarChar, &H0001,2147483647)
	p.Value = request("txtDetails")
	cm.Parameters.Append p

	if trim(request("txtDisplayedID"))="" then
		Set p = cm.CreateParameter("@NewID",adInteger, &H0002)
		cm.Parameters.Append p
	end if					
	
	
	cm.Execute rowsupdated

	if trim(request("txtDisplayedID"))="" then
		strNewID = cm("@NewID")
	end if

	Set cm = Nothing
	
	
	if rowsupdated = 1 then
		'Update Roadmap Order	
		strSuccess = "1"
		set rs = server.CreateObject("ADODB.recordset")
		rs.Open "spListActionRoadmap " & clng(request("cboProject")),cn,adOpenForwardOnly
		if request("cboPriority") = "" then			i=1
			cn.execute "spUpdateActionRoadmapDisplayOrder " & strNewID & "," & i 
			i=2'			Response.Write "<BR>Move to Front<BR>"			Response.write "ID:" & strNewID & ":" & i & "<BR>"
		else
			i=1
		end if
		do while not rs.EOF
			if trim(rs("ID")) = trim(request("cboPriority")) then				cn.execute "spUpdateActionRoadmapDisplayOrder " & rs("ID") & "," & i 
				Response.write "ID:" & rs("ID") & ":" & i & "<BR>"
				i=i+1
				cn.execute "spUpdateActionRoadmapDisplayOrder " & strNewID & "," & i				Response.write "ID:" & strNewID & ":" & i & "<BR>"
				i=i+1 
				if cn.errors.count > 0 then					strsuccess = ""
					exit do				end if
			elseif trim(rs("ID")) <> trim(strNewID) then				cn.execute "spUpdateActionRoadmapDisplayOrder " & rs("ID") & "," & i 
				Response.write "ID:" & rs("ID") & ":" & i & "<BR>"
				i=i+1 
				if cn.errors.count > 0 then					strsuccess = ""
					exit do				end if
			end if
			rs.MoveNext
		loop
		rs.Close		set rs=nothing
	else
		strSuccess = ""
	end if
	
	cn.Close
	set cn=nothing

%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">

</BODY>
</HTML>
