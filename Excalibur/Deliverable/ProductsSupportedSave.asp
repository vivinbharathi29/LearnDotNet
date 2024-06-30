<%@  language="VBScript" %>
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <script src="../Scripts/PulsarPlus.js"></script>
</head>
<script id="clientEventHandlersJS" language="javascript">
<!--

    function window_onload() {
        if (typeof (txtSuccess) != "undefined") {            
            if (txtSuccess.value == "1") {
                
                if (IsFromPulsarPlus()) {
                    window.parent.parent.parent.popupCallBack(1);
                    ClosePulsarPlusPopup();
                }
                else {
                    window.returnValue = txtSuccess.value;
                    window.parent.close();
                }
            }
            else
                document.write("<BR><font size=2 face=verdana>Unable to update supported products.</font>");
        }
        else
            document.write("<BR><font size=2 face=verdana>Unable to update supported products.</font>");
    }


    //-->
</script>

<body language="javascript" onload="return window_onload()">

    <%
	dim strAdd
	dim strRemove
	dim IDLoadedArray
	dim IDSelectedArray
	dim strItem
	dim strSuccess
	dim blnError
	dim cn
	dim cm
	

	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
		
	cn.BeginTrans
	blnError = false


	'Prep Strings
	IDLoadedArray = split(request("txtProductsLoaded"),",")
	IDSelectedArray = split(request("chkSupport"),",")
	
		
	'Find ones to Add
	for each strItem in IDSelectedArray
		if not InArray(IDLoadedArray,strItem) then
			Response.write "<BR>Add:" & strItem  
			
         	set cm = server.CreateObject("ADODB.Command")
			            
            cm.ActiveConnection = cn
            cm.CommandText = "spLinkVersion2Product"
            cm.CommandType = &H0004
	                
            Set p = cm.CreateParameter("@ProductVersionID", 3, &H0001)
			p.Value = clng(strItem)
            cm.Parameters.Append p
	                    
            Set p = cm.CreateParameter("@DeliverableID", 3, &H0001)
			p.Value = clng(request("txtVersionID"))
            cm.Parameters.Append p
		                    
	        cm.Execute recordseffected
                    
            Set cm = Nothing

			if recordseffected = 0 or cn.Errors.count > 0 Then
				blnError = true
				exit for
			End If
			
		end if
	next
		

	'Find ones to remove
	if not blnError then
		for each strItem in IDLoadedArray
			if not InArray(IDSelectedArray,strItem) then
				Response.write "<BR>Remove:" & strItem
				
          		set cm = server.CreateObject("ADODB.Command")
			            
				cm.ActiveConnection = cn
		        cm.CommandText = "spUnLinkVersionFromProduct"
				cm.CommandType = &H0004
                
				Set p = cm.CreateParameter("@ProdID", 3, &H0001)
				p.Value = clng(strItem)
				cm.Parameters.Append p
                    
				Set p = cm.CreateParameter("@DeliverableID", 3, &H0001)
				p.Value = clng(request("txtVersionID"))
				cm.Parameters.Append p
                    
				cm.Execute recordseffected
                    
				Set cm = Nothing

				If recordseffected = 0 or cn.Errors.count > 0 Then
					blnError = true
					exit for
				End If
					
					  
			end if
		next
	end if
	
	
	if blnError then
		cn.RollbackTrans
		strSuccess = "0"
	else
		cn.CommitTrans
		strSuccess = "1"
	end if
	
	
	
	
	
		function InArray(MyArray,strFind)
			dim strElement
			dim blnFound
			
			blnFound = false
			for each strElement in MyArray
				if trim(strElement) = trim(strFind) then
					blnFound = true
					exit for
				end if
			next
			InArray = blnFound
		end function
	
%>


    <input type="text" id="txtSuccess" name="txtSuccess" value="<%=strSuccess%>">
</body>
</html>
