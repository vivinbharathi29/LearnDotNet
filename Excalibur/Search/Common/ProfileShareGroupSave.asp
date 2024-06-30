<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	var OutArray = new Array();
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value == "1")
			{
			OutArray[0] = txtName.value;
			OutArray[1] = txtID.value;
			window.parent.returnValue=OutArray;
			window.parent.close();
			}
        }
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<%

    dim blnErrors
    dim strID
    
    blnErrors = false
    
    strID = request("txtGroupID")
    
	if trim(request("txtName")) = "" or trim(request("chkEmployee") ) = "" then
		Response.Write "Not enough information provide to display this page."	
	else
        
		set cn = server.CreateObject("ADODB.Connection")
	
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
		cn.Open
		set rs = server.CreateObject("ADODB.Recordset")


		'Get User
        dim CurrentUser
		dim CurrentDomain
		dim CurrentUserID
		
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
	
        '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
		if not (rs.EOF and rs.BOF) then
			CurrentUserID = rs("ID") & ""
        else
			Response.Write("Pulsar User Not Found. Please submit a Pulsar Support request for assistance.")
            Response.End()
		end if
		rs.Close
        set rs = nothing


        cn.begintrans

		if trim(request("txtGroupID")) <> "" and trim(request("chkDelete")) = "1" then
		    set cm = server.CreateObject("ADODB.Command")
		    cm.CommandType =  &H0004
		    cm.ActiveConnection = cn
		    cm.CommandText = "spDeleteEmployeeUserSetting"	

    		Set p = cm.CreateParameter("@ID", 3,  &H0001)
	    	p.Value = clng(request("txtGroupID"))
		    cm.Parameters.Append p

    		cm.Execute RowsUpdated
	    	if RowsUpdated <> 1 then
		    	blnErrors = true
    		end if
    		strID=0
	    	set cm = nothing
		elseif trim(request("txtGroupID")) = "" then
		    set cm = server.CreateObject("ADODB.Command")
		    cm.CommandType =  &H0004
		    cm.ActiveConnection = cn
		    cm.CommandText = "spAddEmployeeUserSetting"	

    		Set p = cm.CreateParameter("@EmployeeID", 3,  &H0001)
	    	p.Value = clng(CurrentUserID)
		    cm.Parameters.Append p

    		Set p = cm.CreateParameter("@UserSettingsID", 3,  &H0001)
	    	p.Value = 4
		    cm.Parameters.Append p

		    Set p = cm.CreateParameter("@Value", 200,  &H0001,8000)
		    p.Value = left(request("txtName") & "|" & replace(trim(request("chkEmployee"))," ",""),8000)
		    cm.Parameters.Append p

    		Set p = cm.CreateParameter("@NewID", 3,  &H0002)
		    cm.Parameters.Append p

    		cm.Execute RowsUpdated

            strID = cm("@NewID")
	    	if RowsUpdated <> 1 then
		    	blnErrors = true
    		end if
	    	set cm = nothing

		else
		    set cm = server.CreateObject("ADODB.Command")
		    cm.CommandType =  &H0004
		    cm.ActiveConnection = cn
		    cm.CommandText = "spUpdateEmployeeUserSetting"	

    		Set p = cm.CreateParameter("@ID", 3,  &H0001)
	    	p.Value = clng(request("txtGroupID"))
		    cm.Parameters.Append p

		    Set p = cm.CreateParameter("@Value", 200,  &H0001,8000)
		    p.Value = left(request("txtName") & "|" & replace(trim(request("chkEmployee"))," ",""),8000)
		    cm.Parameters.Append p

    		cm.Execute RowsUpdated
	    	if RowsUpdated <> 1 then
		    	blnErrors = true
    		end if
	    	set cm = nothing
		
		
		end if
		
		if blnErrors then
			strSuccess = "0"
			cn.rollbacktrans
		else
			strSuccess = "1"
			cn.committrans
		end if		

		cn.Close
		set cn = nothing
	end if
%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<INPUT type="hidden" id=txtName name=txtName value="<%=request("txtName")%>">
<INPUT type="hidden" id=txtID name=txtID value="<%=strID%>">
</BODY>
</HTML>
