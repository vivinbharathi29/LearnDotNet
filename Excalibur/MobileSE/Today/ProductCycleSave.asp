<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (txtSuccess.value!="0")
	{
	    if (parent.window.parent.document.getElementById('modal_dialog')) {
	        //save value and return to parent page: ---
	        parent.window.parent.SelectedCyclesResult(txtSuccess.value);
	        parent.window.parent.modalDialog.cancel();
	    } else {
	        window.returnValue = txtSuccess.value;
	        window.close();
	    }
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%

if trim(request("txtID")) = "0" then
    strSuccess = request("txtCycleList") & "|" & replace(request("chkCycle")," ","")
else
	dim strConnect
	dim cn
	dim cm
	dim rowseffected
    dim strSelected
    dim strLoaded
    dim LoadedArray
    dim SelectedArray
    dim strItem
    dim strSuccess
    
    strSuccess = request("txtCycleList")
    
    strSelected = replace(request("chkCycle")," ","")
    strLoaded = replace(request("txtCycleLoaded")," ","")
	LoadedArray = split(strLoaded,",")
	SelectedArray = split(strSelected,",")
	
	strConnect = Session("PDPIMS_ConnectionString")
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = strConnect
	cn.Open

    cn.begintrans
    
    for each strItem in SelectedArray
        if instr("," & strLoaded & ",","," & strItem & ",") < 1 then
            response.write "<BR>Add Item: " & stritem
            
        	set cm = server.CreateObject("ADODB.command")

        	cm.ActiveConnection = cn
        	cm.CommandText = "spLinkProductToProgram"
        	cm.CommandType =  &H0004

        	Set p = cm.CreateParameter("@Program", 3, &H0001)
        	p.Value = clng(stritem)
        	cm.Parameters.Append p

        	Set p = cm.CreateParameter("@Product", 3, &H0001)
        	p.Value = clng(request("txtID"))
        	cm.Parameters.Append p

        	cm.Execute RowsEffected
        	if RowsEffected <> 1 then
        	    strSuccess = "0"
        	    exit for
        	end if
            set cm = nothing
        end if
    next

    if trim(strSuccess) <> "0" then
        for each strItem in LoadedArray
            if instr("," & strSelected & ",","," & strItem & ",") < 1 then
                response.write "<BR>Remove Item: " & stritem

        	    set cm = server.CreateObject("ADODB.command")

        	    cm.ActiveConnection = cn
        	    cm.CommandText = "spUnLinkProductFromProgram"
        	    cm.CommandType =  &H0004

        	    Set p = cm.CreateParameter("@Program", 3, &H0001)
        	    p.Value = clng(stritem)
        	    cm.Parameters.Append p

        	    Set p = cm.CreateParameter("@Product", 3, &H0001)
        	    p.Value = clng(request("txtID"))
        	    cm.Parameters.Append p

        	    cm.Execute RowsEffected
        	    if RowsEffected <> 1 then
        	        strSuccess = "0"
        	        exit for
        	    end if
                set cm = nothing

            end if
        next
    end if
    
    if trim(strSuccess) = "0" then
        cn.rollbacktrans
    else
        cn.committrans        
    end if

'	set cm = server.CreateObject("ADODB.command")
		
'	cm.ActiveConnection = cn
'	cm.CommandText = "spUpdateTodayConfig"
'	cm.CommandType =  &H0004


'	Set p = cm.CreateParameter("@ID", 3, &H0001)
'	p.Value = clng(request("txtUserID"))
'	cm.Parameters.Append p

'	cm.Execute RowsEffected
	
'	if cn.Errors.Count > 1 then
'		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""0"">"
'		Response.Write "<font size=2 face=verdana><b>Unable to save this configuration.</b></font>"
'	else
'		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""1"">"
'	end if
	

    cn.close
    
    set cm = nothing
	set cn = nothing

end if

%>
    <input id="txtSuccess" name="txtSuccess" type="text" value="<%=strSuccess%>">
</BODY>
</HTML>
