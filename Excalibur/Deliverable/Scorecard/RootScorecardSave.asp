<%@ Language=VBScript %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload() {
        if (typeof (txtSuccess) != "undefined") {
            if (txtSuccess.value != "0") {
                window.parent.returnValue = txtSuccess.value;
							var iframeName = parent.window.name;
				    	if (iframeName != '') {
				        parent.window.parent.ClosePopUp();
				    	} else {
                window.parent.close();
            }
        }
    }

    }
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload();">
<%

	dim cn
	dim cm
	dim rs
	dim strSuccess
	Dim RowsUpdated

	set cn = server.CreateObject("ADODB.connection")
	set rs = server.CreateObject("ADODB.recordset")

	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
    if request("txtActionRequest")="1" then
        if trim(request("txtAction")) = trim(request("txtActionTemplate")) then
            strSuccess=request("txtActionStatus") & "&nbsp;"
        else
            strSuccess=request("txtActionStatus") & replace(request("txtAction"),vbcrlf,"<BR>")
        end if
    else
	    strSuccess = "1"
    end if

	cn.BeginTrans

	set cm = server.CreateObject("ADODB.Command")
				            
	cm.ActiveConnection = cn
	cm.CommandText = "spUpdateDeliverableScorecard"
	cm.CommandType = &H0004
	                
	Set p = cm.CreateParameter("@RootID", 3, &H0001)
	p.Value = clng(request("txtID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ExecutiveSummary", 200, &H0001,400)
	if trim(request("txtExecutiveSummary")) = trim(request("txtExecutiveSummaryTemplate")) then
        p.Value = ""
    else
        p.Value = left(request("txtExecutiveSummary"),400)
	end if
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@ExecutiveSummaryStatus", 3, &H0001)
	p.Value = clng(request("txtExecutiveSummaryStatus"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@HPPeopleProcess", 200, &H0001,400)
	if trim(request("txtHPPeopleProcess")) = trim(request("txtHPPeopleProcessTemplate")) then
        p.Value = ""
    else
    	p.Value = left(request("txtHPPeopleProcess"),400)
	end if
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@HPPeopleProcessStatus", 3, &H0001)
	p.Value = clng(request("txtHPPeopleProcessStatus"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@HPEquipment", 200, &H0001,400)
	if trim(request("txtHPEquipment")) = trim(request("txtHPEquipmentTemplate")) then
        p.Value = ""
    else
    	p.Value = left(request("txtHPEquipment"),400)
	end if
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@HPEquipmentStatus ", 3, &H0001)
	p.Value = clng(request("txtHPEquipmentStatus"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@SupplierPeopleProcess", 200, &H0001,400)
	if trim(request("txtSupplierPeopleProcess")) = trim(request("txtSupplierPeopleProcessTemplate")) then
        p.Value = ""
    else
    	p.Value = left(request("txtSupplierPeopleProcess"),400)
	end if
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@SupplierPeopleProcessStatus", 3, &H0001)
	p.Value = clng(request("txtSupplierPeopleProcessStatus"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@SupplierDeliverables", 200, &H0001,400)
	if trim(request("txtSupplierDeliverables")) = trim(request("txtSupplierDeliverablesTemplate")) then
        p.Value = ""
    else
    	p.Value = left(request("txtSupplierDeliverables"),400)
	end if
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@SupplierDeliverablesStatus", 3, &H0001)
	p.Value = clng(request("txtSupplierDeliverablesStatus"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Action", 200, &H0001,800)
	if trim(request("txtAction")) = trim(request("txtActionTemplate")) then
        p.Value = ""
    else
    	p.Value = left(request("txtAction"),800)
	end if
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@ActionStatus", 3, &H0001)
	p.Value = clng(request("txtActionStatus"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ShowOnStatus",  adBoolean, &H0001)
	if trim(request("chkStatusReport")) = "1" then
        p.Value = 1
    else
        p.Value = 0
    end if
	cm.Parameters.Append p
		                    
                            

	cm.Execute recordseffected
	                    
	Set cm = Nothing

	if cn.Errors.count > 0 or recordseffected <> 1 then
		cn.RollbackTrans
        strSuccess = "0"
		Response.Write "<BR><BR><font size=2 face=verdana>Unable to update this scorecard.</font>"
	else
		cn.CommitTrans
	end if

	cn.close
	set cn = nothing

%>
    <textarea style="display:none" id="txtSuccess" rows="2"><%=strSuccess%></textarea>
</BODY>
</HTML>



