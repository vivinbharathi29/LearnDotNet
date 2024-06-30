<%@ Language=VBScript %>

<!-- #include file = "../../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload() {
        if (typeof (txtSuccess) != "undefined") {
            var sRowID = document.getElementById("RowID").value;
            //close window
            if (IsFromPulsarPlus()) {
                window.parent.parent.parent.parent.popupCallBack(sRowID);
                ClosePulsarPlusPopup();
            }
            else {
                if (sRowID == "")
                    window.parent.Close("");
                else {
                    window.parent.Close(sRowID);
                }
            }
        }
    }



//-->
</SCRIPT>

</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%
	dim cn
	dim cm
	dim strSuccess
	dim strID
	dim IDArray
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	dim CurrentDomain
	dim CurrentUser
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
	
	if not (rs.EOF and rs.BOF) then
    	CurrentUserID = rs("ID") & ""
    else
	    CurrentUserID = 0
	end if
	rs.Close
    set rs = nothing

	cn.BeginTrans
	if request("txtType") = "1" then
		if not(isnumeric(request("txtSub")) and len(request("txtSub"))=6) then
			strSuccess = request("txtSub")
		else
			strSuccess = request("txtSub") & "-XXX"
		end if
	else 
		strSuccess = request("txtDash")
	end if       
      
    IDArray = split(request("txtID"),",")
    dim arrDvR
    for each strID in IDArray
        
        arrDvR = split(strID,"_")

	    set cm = server.CreateObject("ADODB.Command")
	    cm.CommandType =  &H0004
	    cm.ActiveConnection = cn

	    if request("txtEngCoordinator") = "1" and request("txtSvcCoordinator") = "1" and arrDvR(1) > 0 then
	        cm.CommandText = "spUpdateSubassemblyNumberRelease"	
        elseif request("txtEngCoordinator") = "1" and request("txtSvcCoordinator") = "1" and arrDvR(1) = 0 then
            cm.CommandText = "spUpdateSubassemblyNumber"
        elseif request("txtEngCoordinator") = "1" and arrDvR(1) > 0 then
	        cm.CommandText = "spUpdateSubassemblyNumberEngineeringRelease"
        elseif request("txtEngCoordinator") = "1" and arrDvR(1) = 0 then
	        cm.CommandText = "spUpdateSubassemblyNumberEngineering"
        elseif request("txtSvcCoordinator") = "1" and arrDvR(1) > 0 then
	        cm.CommandText = "spUpdateSubassemblyNumberServiceRelease"
        elseif request("txtSvcCoordinator") = "1" and arrDvR(1) = 0 then
	        cm.CommandText = "spUpdateSubassemblyNumberService"	
	    end if

	    Set p = cm.CreateParameter("@RootID", 3,  &H0001)
	    p.Value = clng(request("txtRootID"))
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@ID", 3,  &H0001)
	    p.Value = clng(trim(arrDvR(0)))
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@UserID", 3,  &H0001)
	    p.Value = clng(currentuserid)
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@UserName", 200,  &H0001,80)
	    p.Value = left(CurrentDomain & "_" & Currentuser ,80)
	    cm.Parameters.Append p

        if request("txtEngCoordinator") = "1" then
	        Set p = cm.CreateParameter("@Sub", 200,  &H0001,6)
	        p.Value = UCase(left(request("txtSub"),6))
	        cm.Parameters.Append p

    	    Set p = cm.CreateParameter("@Spin", 200,  &H0001,3)
    	    p.Value = UCase(left(request("txtDash"),3))
    	    cm.Parameters.Append p

	        Set p = cm.CreateParameter("@FullNumber", 200,  &H0001,10)
	        if request("txtType") = "1" then
    		    p.Value = ""
    	    elseif request("txtSub") <> "" and request("txtDash") <> "" then
		        p.Value = UCase(left(request("txtSub") & "-" & request("txtDash"),10))
	        else
    		    p.value = ""
    	    end if
    	    cm.Parameters.Append p
        end if
        
        if request("txtSvcCoordinator") = "1" then
	        Set p = cm.CreateParameter("@ServiceSub", 200,  &H0001,6)
	        p.Value = UCase(left(request("txtServiceSub"),6))
	        cm.Parameters.Append p

	        Set p = cm.CreateParameter("@ServiceSpin", 200,  &H0001,3)
	        p.Value = UCase(left(request("txtServiceDash"),3))
	        cm.Parameters.Append p

	        Set p = cm.CreateParameter("@ServiceFullNumber", 200,  &H0001,10)
            if request("txtServiceSub") <> "" and request("txtServiceDash") <> "" then
		        p.Value = UCase(left(request("txtServiceSub") & "-" & request("txtServiceDash"),10))
	        else
		        p.value = ""
	        end if
	        cm.Parameters.Append p
        end if	
    	
        if arrDvR(1) > 0 then
            Set p = cm.CreateParameter("@ReleaseID", 3,  &H0001)
	        p.Value = clng(trim(arrDvR(1)))
	        cm.Parameters.Append p
        end if

	    cm.Execute rowschanged
    	
	    set cm=nothing

	    if  cn.Errors.count > 0 then 'rowschanged = 0 or
		    Response.Write "DASH:" & left(request("txtDash"),3) & "<BR>SUB:" & left(request("txtSub"),6) & "<BR>ID:" & clng(request("txtID")) & "<BR>RootID:" & clng(request("txtRootID")) & "<BR>RC:" & rowschanged & "<BR>EC:" & cn.Errors.count
		    strSuccess = "0"
		    exit for
	    end if	
    next

    set cm = server.CreateObject("ADODB.Command")
	cm.ActiveConnection = cn
	if Request("ID") <> "" then
	
		cn.execute "spLogDeliverableUpdate " & clng(CurrentUserID) & "," & clng(Request("ID")) & ",0"

		cm.CommandText = "spUpdateDelRootWeb2"
		cm.CommandType =  &H0004
		
		Set p = cm.CreateParameter("@ID", 3,  &H0001)
		p.Value = Request("ID")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Name", 200, &H0001, 120)
		p.value = left(trim(replace(request("txtDelName"),chr(150),"-")),120)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@NameElements", 200, &H0001, 1000)
		p.value = trim(request("strNameElements"))
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@Name2", 200, &H0001, 120)
		p.value = left(trim(replace(request("txtDelName2"),chr(150),"-")),120)
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@Name3", 200, &H0001, 120)
		p.value = left(trim(replace(request("txtDelName3"),chr(150),"-")),120)
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@Name4", 200, &H0001, 120)
		p.value = left(trim(replace(request("txtDelName4"),chr(150),"-")),120)
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@Name5", 200, &H0001, 120)
		p.value = left(trim(replace(request("txtDelName5"),chr(150),"-")),120)
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@Name6", 200, &H0001, 120)
		p.value = left(trim(replace(request("txtDelName6"),chr(150),"-")),120)
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@Name7", 200, &H0001, 120)
		p.value = left(trim(replace(request("txtDelName7"),chr(150),"-")),120)
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@Name8", 200, &H0001, 120)
		p.value = left(trim(replace(request("txtDelName8"),chr(150),"-")),120)
		cm.Parameters.Append p
		
		cm.Execute rowschanged

		if rowschanged <> 1 then
			Response.Write "<BR>Error while saving root.  Verify the Name is unique and/or check Database table<BR>" & Request.form & Request.QueryString
			FoundErrors = true
		end if
		NewID = request("ID")	
	end if	    
	
	if strSuccess = "0" then
		cn.RollbackTrans
	else
		cn.CommitTrans
	end if

	cn.Close
	set cn = nothing


%>
<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<input style="display: none" type="text" id="RowID" name="RowID" value="<%=Request("RowID")%>" />

</BODY>
</HTML>