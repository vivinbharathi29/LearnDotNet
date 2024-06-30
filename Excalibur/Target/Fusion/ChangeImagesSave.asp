<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <script src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload(pulsarplusDivId) {
	if (typeof(txtSuccess) != "undefined"){
	    if (txtSuccess.value != "0") {
	        //close window
	        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
	            parent.window.parent.closeExternalPopup();
	            parent.window.parent.reloadFromPopUp('Deliverables');
	        }
	        else if (IsFromPulsarPlus()) {
	            window.parent.parent.parent.popupCallBack(1);
	            ClosePulsarPlusPopup();
	        }
	        else {
	            if (parent.window.parent.document.getElementById('modal_dialog')) {
	                var strResult = parent.window.parent.modalDialog.getArgument('change_image_cell');
	                //save value and return to parent page: ---
	                if (strResult == 'IMGCell') {
	                    parent.window.parent.ChangeImageResult(txtSuccess.value);
	                } else {
	                    parent.window.parent.ModImagesResult(txtSuccess.value);
	                }
	                parent.window.parent.modalDialog.cancel();
	            } else {
	                window.returnValue = txtSuccess.value;
	                window.close();
	            }
	        }
		}
	//	else
	//		document.write ("Unable to update Images.  An unexpected error occurred.");	
	}
	//else
	//	{
	//	document.write ("Unable to update Images.  An unexpected error occurred.");
	//	}

}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">

<%
	dim i
	dim cn
	dim rs
	dim cm
	dim strSuccess
	dim RowsEffected
	dim blnFailed
	
	
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.recordset")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	'Get User
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

	set cm=nothing
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
	else
		CurrentUserID = ""
	end if
	
	rs.Close

	cn.BeginTrans
	
	blnFailed = false

    response.write "Version: " & request("txtVersionID") & "<br>"
    response.write "Product: " & request("txtProductID") & "<br>"

    if request("txtVersionID") <> "" then
	    set cm = server.CreateObject("ADODB.Command")
	
	    cm.ActiveConnection = cn
	    cm.CommandText = "spUpdateImageActuals"
	    cm.CommandType = &H0004
	      
	    Set p = cm.CreateParameter("@ProdID",adInteger, &H0001)
	    p.Value = request("txtProductID")
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@VerisonID",adInteger, &H0001)
	    p.Value = request("txtVersionID")
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@ImageSummary",200, &H0001,8000)
	    p.Value = left(request("txtSummary"),8000)
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@ImageList", 201, &H0001, 2147483647)
	    if request("chkAllChecked") = "on" and request("txtPatch") <> "1" then
		    p.value= request("txtLangList")
	    else
		    if request("chkImage") <> "" then
			    p.value = request("chkImage") & "," & request("txtLangList")
		    else
			    p.value = "0," & request("txtLangList")
		    end if
	    end if
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@RestoreImageList", 200, &H0001, 2147483647)
	    if request("chkRestoreImage") = ""  then
            p.value = ""
        else
            if request("txtRestoreLangList") <> "" then
			    p.value = request("chkRestoreImage") & "," & request("txtRestoreLangList")
		    else
			    p.value = request("chkRestoreImage") 
		    end if
	    end if
	    cm.Parameters.Append p
		
        cm.Execute RowsEffected
	    Set cm = Nothing
	    if rowseffected <> 1 then
		    blnFailed = true
	    end if
    end if

	if (not blnFailed) and (trim(request("optScope")) = "2" or request("txtVersionID") = "" ) then
		set cm = server.CreateObject("ADODB.Command")
	
	    cm.ActiveConnection = cn
	    cm.CommandText = "spUpdateImageDefaults"
	    cm.CommandType = &H0004
	       
		Set p = cm.CreateParameter("@ProdID",adInteger, &H0001)
		p.Value = request("txtProductID")
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@RootID",adInteger, &H0001)
		p.Value = request("txtRootID")
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@ImageSummary",200, &H0001,8000)
		p.Value = left(request("txtSummary"),8000)
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@ImageList", 201, &H0001, 2147483647)
		if request("chkAllChecked") = "on" then
			p.value=request("txtLangList")
		else
			if request("chkImage") <> "" then
				p.value = request("chkImage") & "," & request("txtLangList")
			else
				p.value = "0," & request("txtLangList")
			end if
		end if
		cm.Parameters.Append p
	
	    Set p = cm.CreateParameter("@RestoreImageList", 200, &H0001, 2147483647)
	    if request("chkRestoreImage") = ""  then
            p.value = ""
        else
            if request("txtRestoreLangList") <> "" then
	    		p.value = request("chkRestoreImage") & "," & request("txtRestoreLangList")
	    	else
	    		p.value = request("chkRestoreImage") 
	    	end if
	    end if
	    cm.Parameters.Append p

	    cm.Execute RowsEffected
		Set cm = Nothing

		if rowseffected <> 1 then
			blnFailed = true
		end if

	end if
    if (not blnfailed) and trim(currentuserid) <> "" and request("txtVersionID") <> "" then
        cn.execute "spLogAction2 " & clng(currentuserid) & ",40," & clng(request("txtProductID")) & "," & clng(request("txtVersionID")) & ",''", rowseffected
        if rowseffected <> 1 then
            blnfailed = true
        end if
    end if
	
	if blnFailed then
		strSuccess = "0"
		cn.RollbackTrans
	else
		strSuccess = left(request("txtSummary"),80)
		cn.CommitTrans
	end if
		
	set rs=nothing
	set cn=nothing
	
%>

<INPUT type="text" style="display:" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>
