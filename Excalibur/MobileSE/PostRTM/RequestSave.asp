<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <script src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value == "1")
			{
		    window.returnValue = 1;
		    if (IsFromPulsarPlus()) {
		        if (GetQueryStringValue("component") == "RequestSustaining") {
		            window.parent.parent.parent.RequestComponentCallback(txtSuccess.value);
		        }
		        else {
		            window.parent.parent.parent.SustainingProductSupportCallback(txtSuccess.value);
		        }
		        ClosePulsarPlusPopup();
		    }
		    else {
		        window.parent.close();
		    }
			}
	//	else
	//		document.write ("<BR><font size=2 face=verdana>Unable to update the record because strSuccess 0.</font>");
		}
	//else
	//	document.write ("<BR><font size=2 face=verdana>Unable to update the record because strSuccess is not defined.</font>");
}

//-->
</SCRIPT>

</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%
	dim cn
	dim rs
	dim strSuccess
	dim ProductArray
	dim ProductStatusArray
	dim StopWatchingRoot
	dim StopWatchingVersion
	dim WatchingProduct
	dim RootID
	dim VersionID
		
	strSuccess = "1"
	StopWatchingRoot=0
	StopWatchingVersion=0

	'kz 10/12/04. Don't check "stop watch" any more after added more status 
	RootID = request("cboRoot")
	'if instr(RootID,"*") >0 then
	'	RootID = left(RootID, instr(RootID,"*")-1)
	'	StopWatchingRoot = 1
	'end if
	VersionID = request("cboVersion")
	'if instr(VersionID,"*") >0 then
	'	VersionID = left(VersionID, instr(VersionID,"*")-1)
	'	StopWatchingVersion = 1
	'end if

	'if request("cboRoot") = "" then
	'	strSuccess = "0"
	'end if	
		
	'if request("cboVersion") = "" then
	'	strSuccess = "0"
	'end if	

	'if request("txtProductIDs") = "" then
	'	strSuccess = "0"
	'end if	
	
	'if request("cboProductStatus") = "" then
	'	strSuccess = "0"
	'end if	

	'if request("cboVersionStatus") = "" then
	'	strSuccess = "0"
	'end if	
	
	ProductArray = Split(request("txtProductIDs"), "*", -1, 1)
	ProductStatusArray = Split(request("cboProductStatus"), ",", -1, 1)
	
	set cn = server.CreateObject("ADODB.Connection")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open

	cn.BeginTrans
	set cm = server.CreateObject("ADODB.Command")
	'&H0004 means stored procedure
	cm.CommandType =  &H0004
	cm.ActiveConnection = cn
		
	'Set the Root deliverable's PostRTM field to be 1 once it is selected.
	'and set it to 0 if stop watching is selected
	cm.CommandText = "spUpdatePostRTMRoot"		
	Set p = cm.CreateParameter("@RootID", 3,  &H0001)
	p.Value = RootID

	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Status", 3,  &H0001)
	'if StopWatchingRoot = 1 then
	'	p.Value = 0
	'else
	p.Value = 1
	'end if
	cm.Parameters.Append p

	cm.Execute rowschanged
	if cn.Errors.count > 0 then
		strSuccess = "0"
	end if	

	if strSuccess = "0" then
		cn.RollbackTrans
	else
		cn.CommitTrans
	end if

	Set cm=nothing
	Set p=nothing

	set rs = server.CreateObject("ADODB.recordset")
	set rs2 = server.CreateObject("ADODB.recordset")

	WatchingProduct = 1
'	if StopWatchingRoot = 1 then
'		WatchingProduct = 0
'		'check to see if all the version RTM field is set to 0 or not
'		'If it is all set to 0, we should stop watching the products when RootID is set to stop watching
'		rs.open "spListDeliverableVersions " & RootID ,cn,adOpenForwardOnly
'		do while not rs.EOF
'			rs2.open "spGetDeliverableVersionProperties " & rs("VersionID"),cn,adOpenForwardOnly
'			if trim(rs2("PostRTMStatus")) <> "0" then
'				WatchingProduct = 1
'			end if
'			rs2.close
'			rs.MoveNext
'		loop
'		rs.Close
'	end if

	' If Stop Watch root is set and no Version's RTM status is set, change all
	' the products RTM status to recommnded
'	if StopWatchingRoot = 1 and WatchingProduct=0 and StopWatchingVersion <> 1 then
'		cn.BeginTrans
'		for i = lbound(ProductArray) to (ubound(ProductArray)-1)
'			' set all the Products's PostRTM field to 0
'			set cm = server.CreateObject("ADODB.Command")
'			'&H0004 means stored procedure
'			cm.CommandType =  &H0004
'			cm.ActiveConnection = cn
'			cm.CommandText = "spUpdatePostRTMRootProduct"	
'
'			Set p = cm.CreateParameter("@RootID", 3,  &H0001)
'			p.Value = RootID
'			cm.Parameters.Append p
'
'			Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
'			p.Value = trim(ProductArray(i))
'			cm.Parameters.Append p
'
'			Set p = cm.CreateParameter("@StatusID", 3,  &H0001)
'			p.Value = 0
'			cm.Parameters.Append p
'
'			cm.Execute rowschanged
'			if rowschanged <> 1 then
'				strSuccess = "0"
'				Exit for
'			end if	
'			Set cm=nothing
'			Set p=nothing
'		next
'		if strSuccess = "0" then
'			cn.RollbackTrans
'		else
'			cn.CommitTrans
'		end if
'	end if
	' if stop watch root and stop watch version both not set
	' set all the Products's PostRTM field to the status that is selected
	if StopWatchingRoot = 0 and StopWatchingVersion = 0 then
		cn.BeginTrans
		for i = lbound(ProductArray) to (ubound(ProductArray)-1)
			set cm = server.CreateObject("ADODB.Command")
			'&H0004 means stored procedure
			cm.CommandType =  &H0004
			cm.ActiveConnection = cn
			cm.CommandText = "spUpdatePostRTMRootProduct"	
	
			Set p = cm.CreateParameter("@RootID", 3,  &H0001)
			p.Value = RootID
			cm.Parameters.Append p
	
			Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
				p.Value = ProductArray(i)
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@StatusID", 3,  &H0001)
			p.Value = ProductStatusArray(i)
			cm.Parameters.Append p

			cm.Execute rowschanged
			if rowschanged <> 1 then
				strSuccess = "0"
				Exit for
			end if	
			Set cm=nothing
			Set p=nothing
		next
		if strSuccess = "0" then
			cn.RollbackTrans
		else
			cn.CommitTrans
		end if
	end if
	'Set the Status in DeliverableVersion to whatever was selected if a Version is selected
	if VersionID <> 0 then 'and StopWatchingRoot <> 1 and StopWatchingVersion <> 1 then
		cn.BeginTrans
		set cm = server.CreateObject("ADODB.Command")
		'&H0004 means stored procedure
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn
		cm.CommandText = "spUpdatePostRTMDeliverableVersion"	

		Set p = cm.CreateParameter("@VersionID", 3,  &H0001)
		p.Value = VersionID
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Status", 3,  &H0001)
		p.Value = request("cboVersionStatus")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@TargetDate", 135,  &H0001)
		if trim(request("txtTargetDate")) = "" then
			p.Value = null
		else
			p.Value = request("txtTargetDate")
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Comments", 200,  &H0001,120)
		if trim(request("txtComments")) = "" then
			p.Value = null
		else
			p.Value = left(request("txtComments"),120)
		end if
		cm.Parameters.Append p

		cm.Execute rowschanged
		if rowschanged <> 1 then
			strSuccess = "0"
		end if	
		Set cm=nothing
		Set p=nothing

		if cn.Errors.count > 0 then
			strSuccess = "0"
		end if	

		if strSuccess = "0" then
			cn.RollbackTrans
		else
			cn.CommitTrans
		end if		
	'if StopWatchingVersion is 1, set the PostRTMstatus on this version to be 0
'	elseif StopWatchingVersion = 1 then
'		cn.BeginTrans
'		set cm = server.CreateObject("ADODB.Command")
'		'&H0004 means stored procedure
'		cm.CommandType =  &H0004
'		cm.ActiveConnection = cn
'		cm.CommandText = "spUpdatePostRTMDeliverableVersion"	
'
'		Set p = cm.CreateParameter("@VersionID", 3,  &H0001)
'		p.Value = VersionID
'		cm.Parameters.Append p
'
'		Set p = cm.CreateParameter("@Status", 3,  &H0001)
'		p.Value = 0
'		cm.Parameters.Append p
'
'		cm.Execute rowschanged
'		if rowschanged <> 1 then
'			strSuccess = "0"
'		end if	
'		Set cm=nothing
'		Set p=nothing
'
'		if cn.Errors.count > 0 then
'			strSuccess = "0"
'		end if	
'
'		if strSuccess = "0" then
'			cn.RollbackTrans
'		else
'			cn.CommitTrans
'		end if
	end if		

	cn.Close
	set rs = nothing
	set rs2 = nothing
	set cn = nothing

%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>
