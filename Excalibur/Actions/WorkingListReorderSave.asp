<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
	    if (txtSuccess.value == "1") {
	        if (IsFromPulsarPlus()) {
	            window.parent.parent.parent.WorkingListReorderCallBack(txtSuccess.value);
	            ClosePulsarPlusPopup();
	        } else {
	            window.returnValue = 1;
	            window.parent.close();
	        }
	    }
	    else
	        document.write("<BR><font size=2 face=verdana>Unable to update action item order.</font>");
		}
	else
		document.write ("<BR><font size=2 face=verdana>Unable to update action item order.</font>");
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%

	dim i,j, LastIndex
	dim IDArray
	dim ValueArray
	dim OutputArray
	dim cn
	dim MaxValue

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	
	IDArray = split(request("txtIDList"),",")
	OutputArray = split(request("txtIDList"),",")
	ValueArray = split(request("txtValueList"),",")
	
    if trim(request("txtReportOption")) = "2" then
        MaxValue = 0
        for i = lbound(ValueArray) to ubound(ValueArray)
	        if clng(MaxValue) < clng(ValueArray(i)) then
	            MaxValue =  clng(ValueArray(i))
	        end if
	    next
        if Maxvalue > 0 then
            for i = lbound(ValueArray) to ubound(ValueArray)
	            if clng(ValueArray(i)) = 0 then
	                MaxValue = MaxValue + 1
	                ValueArray(i) = clng(MaxValue)
	            end if
	        next
        	
        	for i = 0 to ubound(OutputArray) 
		        cn.Execute "UpdateDeliverableActionDisplayOrder " &  OutputArray(i) & "," & ValueArray(i)
	            'response.write "<BR>" & OutputArray(i) & "," & trim(ValueArray(i))
	        next
	        
        end if

    else
        for i = lbound(OutputArray) to ubound(OutputArray) 
	        OutputArray(i) = ""
        next
        for i = lbound(ValueArray) to ubound(ValueArray)
	        if ValueArray(i) <> 0 then
		        OutputArray(ValueArray(i)-1) = IDArray(i)
	        end if
        next

    	LastIndex=lbound(OutputArray)
    	for i = lbound(ValueArray) to ubound(ValueArray) 
		    if ValueArray(i) = 0 then
    			for j = LastIndex to ubound(OutputArray)
				    if OutputArray(j) = "" then
    					OutputArray(j) = IDArray(i)
					    LastIndex=j+1
					    exit for
				    end if
			    next
		    end if
	    next
    
    	for i = 0 to ubound(OutputArray) 
	        response.write OutputArray(i) & "-" & i+1 & "<BR>"
		    cn.Execute "UpdateDeliverableActionDisplayOrder " &  OutputArray(i) & "," & i+1
	    next
	end if
	
	cn.close
	set cn = nothing
%>
<INPUT type="text" id=txtSuccess name=txtSuccess value="1">
</BODY>
</HTML>
