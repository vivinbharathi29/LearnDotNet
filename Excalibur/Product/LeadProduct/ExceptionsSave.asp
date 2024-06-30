<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../../includes/noaccess.inc" -->

<HTML>
<HEAD>
<TITLE>Add Lead Product Synchronization Exceptions</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="../../Scripts/PulsarPlus.js"></script>
<script src="../../Scripts/Pulsar2.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
    if (typeof (txtSuccess) != "undefined") {
        if (txtSuccess.value == "1") {
            window.returnValue = 1;
            if (isFromPulsar2()) {
                closePulsar2Popup(true);
            }
            else if (IsFromPulsarPlus()) {
                window.parent.parent.parent.LeadProductSynchronizationCallback(txtSuccess.value);
                ClosePulsarPlusPopup();
            }
            else {
                window.close();
            }
            //window.close();
        }
    }
}

//-->
</SCRIPT>
</HEAD>


<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">
<BR>
<table width=100%><TR><TD align=center>
<font face=verdana size =2>Saving Exceptions.  Please wait...</font>
</td></tr></table>

<%

	dim strVersionID
	dim strProductID
	dim cn
	dim rs
	dim CurrentDomain
	dim CurrentUserID
	dim CurrentUser
	dim cm
	dim VersionArray

	if request("txtProductID") = "" or (request("optType") <> "1" and request("optType") <> "2"and request("optType") <> "3") or request("txtVersions") = "" or request("txtRootID") = ""    then
		Response.Write "<BR><font face=verdana size=2>Not enough information supplied</font><BR>" 
		strSuccess = 0
	else


		set cn = server.CreateObject("ADODB.Connection")
		set rs = server.CreateObject("ADODB.recordset")
	
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
		cn.Open


		'Get User
		CurrentUser = lcase(Session("LoggedInUser"))

		if instr(currentuser,"\") > 0 then
			CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
			Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
		end if

		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetUserInfo"
	

		Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
		p.Value = Currentuser
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
		p.Value = CurrentDomain
		cm.Parameters.Append p

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 

		set cm=nothing
	
		if not (rs.EOF and rs.BOF) then
			CurrentUserID = rs("ID") 
		end if
		rs.Close
	
		cn.BeginTrans
		
		if trim(request("optType")) = "1" then
			strProductID = clng(request("txtProductID")	)
			strRootID = clng(request("txtRootID"))
	
			set cm = server.CreateObject("ADODB.Command")
		
			cm.ActiveConnection = cn
			cm.CommandText = "spUpdateLeadProductRootExceptions"
			cm.CommandType = &H0004
							
			Set p = cm.CreateParameter("@ProductVersionID",adInteger, &H0001)
			p.Value = clng(strProductID)
			cm.Parameters.Append p
				
			Set p = cm.CreateParameter("@DeliverableRootID",adInteger, &H0001)
			p.Value = clng(strRootID)
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@SyncDistribution",adBoolean, &H0001)
			if request("chkDistributionAll") = "on" then 'These values are backward on purpose
				p.Value = 0
			else
				p.Value = 1
			end if
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@SyncImages",adBoolean, &H0001)
			if request("chkImagesAll") = "on" then 'These values are backward on purpose
				p.Value = 0
			else
				p.Value = 1
			end if
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@SyncNotes",adBoolean, &H0001)
			if request("chkTargetNotesAll") = "on" then 'These values are backward on purpose
				p.Value = 0
			else
				p.Value = 1
			end if
			cm.Parameters.Append p
						
			Set p = cm.CreateParameter("@Comments", adVarChar, &H0001,8000)
			p.Value = left(request("txtComments"),8000)
			cm.Parameters.Append p

            Set p = cm.CreateParameter("@ReleaseID",adInteger, &H0001)
			p.Value = clng(request("txtReleaseID"))
			cm.Parameters.Append p

			cm.Execute
			Set cm = Nothing		

			if cn.Errors.count > 0 then
				strSuccess = "0"
				cn.RollbackTrans
			else
				strSuccess="1"
			end if
		
		elseif trim(request("optType")) = "2" then
		
			strProductID = clng(request("txtProductID")	)
			strRootID = clng(request("txtRootID"))
		
			VersionArray = split(request("txtVersions"),",")
			
			'Reset all usages to 1 for all versions of the Root before updating the version properties
			set cm = server.CreateObject("ADODB.Command")
		
			cm.ActiveConnection = cn
			cm.CommandText = "spResetLeadProductRootExceptions"
			cm.CommandType = &H0004
							
			Set p = cm.CreateParameter("@ProductVersionID",adInteger, &H0001)
			p.Value = clng(strProductID)
			cm.Parameters.Append p
				
			Set p = cm.CreateParameter("@DeliverableRootID",adInteger, &H0001)
			p.Value = clng(strRootID)
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@Comments", adVarChar, &H0001,8000)
			p.Value = left(request("txtComments"),8000)
			cm.Parameters.Append p

            Set p = cm.CreateParameter("@ReleaseID",adInteger, &H0001)
			p.Value = clng(request("txtReleaseID"))
			cm.Parameters.Append p

			cm.Execute
			Set cm = Nothing		

			if cn.Errors.count > 0 then
				strSuccess = "0"
				cn.RollbackTrans
			else
				strSuccess="1"
			end if



			'process Version exceptions
			if strSuccess="1" and ubound(VersionArray) <> -1 then
				for each strValue in VersionArray
					if not isnumeric(trim(strValue)) then
						Response.Write "<BR>InvalidID<BR>"
						Response.write "<BR>" & request("txtVersions") & "</BR>"
						strSuccess = "0"
						exit for
					else
						set cm = server.CreateObject("ADODB.Command")
			
						cm.ActiveConnection = cn
						cm.CommandText = "spAddLeadProductVersionExceptions"
						cm.CommandType = &H0004
								
						Set p = cm.CreateParameter("@ProductVersionID",adInteger, &H0001)
						p.Value = clng(strProductID)
						cm.Parameters.Append p
					
						Set p = cm.CreateParameter("@DeliverableVersionID",adInteger, &H0001)
						p.Value = clng(strValue)
						cm.Parameters.Append p
	
						Set p = cm.CreateParameter("@SyncDistribution",adBoolean, &H0001)
						if instr(", " & request("chkDistribution") & ",",", " & trim(strValue) & ",") > 0 then
							p.Value = 0
						else
							p.Value = 1
						end if
						cm.Parameters.Append p

						Set p = cm.CreateParameter("@SyncImages",adBoolean, &H0001)
						if instr(", " & request("chkImages") & ",",", " & trim(strValue) & ",") > 0 then
							p.Value = 0
						else
							p.Value = 1
						end if
						cm.Parameters.Append p

						Set p = cm.CreateParameter("@SyncNotes",adBoolean, &H0001)
						if instr(", " & request("chkTargetNotes") & ",",", " & trim(strValue) & ",") > 0 then
							p.Value = 0
						else
							p.Value = 1
						end if
						cm.Parameters.Append p
						
						cm.Execute
						Set cm = Nothing
		
						if cn.Errors.count > 0 then
							strSuccess = "0"
							cn.RollbackTrans
							exit for
						else
							strSuccess = "1"
						end if
				
					end if
				next
			end if
		elseif trim(request("optType")) = "3" then
		
			strProductID = clng(request("txtProductID")	)
			strRootID = clng(request("txtRootID"))
		
			'Reset all usages to 1 for all versions of the Root before updating the version properties
			set cm = server.CreateObject("ADODB.Command")
		
			cm.ActiveConnection = cn
			cm.CommandText = "spResetLeadProductRootExceptions"
			cm.CommandType = &H0004
							
			Set p = cm.CreateParameter("@ProductVersionID",adInteger, &H0001)
			p.Value = clng(strProductID)
			cm.Parameters.Append p
				
			Set p = cm.CreateParameter("@DeliverableRootID",adInteger, &H0001)
			p.Value = clng(strRootID)
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@Comments", adVarChar, &H0001,8000)
			p.Value = "" 'left(request("txtComments"),8000)
			cm.Parameters.Append p

			cm.Execute
			Set cm = Nothing		

			if cn.Errors.count > 0 then
				strSuccess = "0"
				cn.RollbackTrans
			else
				strSuccess="1"
			end if
			
		end if
		
		if strSuccess = "1" then
			cn.CommitTrans
		end if
					            
		set rs = nothing
		set cn = nothing


	end if
%>

<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>


