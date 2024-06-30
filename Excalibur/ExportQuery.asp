<%@ Language=VBScript %>

<!-- #include file = "includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script type="text/javascript" src="includes/client/jquery.min.js"></script>
<script type="text/javascript" src="includes/client/json2.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	var OutArray = new Array();
	//if (txtSuccess.value!="0")
	//	{
		OutArray[0]= txtOut1.value;
		OutArray[1]= txtOut2.value;
		OutArray[2] = txtOut3.value;

        //save array value and return to parent page: ---
		parent.modalDialog.passArgument(JSON.stringify(OutArray), 'export_query_array');
        
        //close window
		if (window.location != window.parent.location) {
		    parent.modalDialog.cancel();
		} else {
		    window.close();
		}

		//window.returnValue = OutArray;
		//window.close();
	//	}
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()" bgcolor=Ivory>
<Font face=verdana size=2><b><br>&nbsp;&nbsp;Accessing the Database. Please wait....</b></font>
<%
	dim strOut1
	dim strOut2
	dim strOut3	
	dim cn
	dim rs
	dim p 
	dim cm
	dim RowsEffected
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	
	Select Case request("Type")
	case "1"
		rs.Open "spGetExcelProfile " & clng(request("ID")),cn,adOpenForwardOnly	
		
		if rs.EOF and rs.BOF then
			strOut1 = ""
			strOut2 = ""
			strOut3 = ""
		else
			strOut1 = rs("ExcelExportProjects") & ""
			strOut2 = rs("ExcelExportColumns") & ""
			strOut3 = rs("ExcelExportHeader") & ""
		end if
		rs.Close
	case "2"
		strOut1 = ""
		strOut2 = ""
		strOut3 = ""

		cn.BeginTrans
		
		set cm = server.CreateObject("ADODB.command")
		
		cm.ActiveConnection = cn
		cm.CommandText = "spRenameExcelProfile"
		cm.CommandType =  &H0004
		
		set p =  cm.CreateParameter("@ID", 3, &H0001)
		p.value =  request("ID")
		cm.Parameters.Append p
	            
		Set p = cm.CreateParameter("@Name", 200, &H0001, 50)
		p.Value = left(Request("NewName"),50)
		cm.Parameters.Append p

		cm.Execute RowsEffected
				
		if cn.Errors.Count > 1 or RowsEffected <> 1 then
			cn.RollbackTrans
			strOut1 = "0"
		else
			cn.CommitTrans
			strOut1 = "1"
		end if
		
		set cm = nothing
	case "3"
		strOut1 = ""
		strOut2 = ""
		strOut3 = ""

		cn.BeginTrans
		
		set cm = server.CreateObject("ADODB.command")
		
		cm.ActiveConnection = cn
		cm.CommandText = "spDeleteExcelProfile"
		cm.CommandType =  &H0004
		
		set p =  cm.CreateParameter("@ID", 3, &H0001)
		p.value =  request("ID")
		cm.Parameters.Append p
	            
		cm.Execute RowsEffected
				
		if cn.Errors.Count > 1 or RowsEffected <> 1 then
			cn.RollbackTrans
			strOut1 = "0"
		else
			cn.CommitTrans
			strOut1 = "1"
		end if
		
		set cm = nothing
			
	case "4"
		strOut1 = ""
		strOut2 = ""
		strOut3 = ""
		
		cn.BeginTrans
		
		set cm = server.CreateObject("ADODB.command")
		
		cm.ActiveConnection = cn
		cm.CommandText = "spAddExcelProfile"
		cm.CommandType =  &H0004
		
		set p =  cm.CreateParameter("@EmployeeID", 3, &H0001)
		p.value =  request("EmployeeID")
		cm.Parameters.Append p
	            
		Set p = cm.CreateParameter("@Name", 200, &H0001, 50)
		p.Value = left(Request("Name"),50)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Products", 200, &H0001, 8000)
		p.Value = left(Request("Products"),8000)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Columns", 200, &H0001, 8000)
		p.Value = left(Request("Columns"),8000)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Header", 16, &H0001)
		if isnumeric(Request("Header")) then
			p.Value = clng(Request("Header"))
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ActionType", 16, &H0001)
		if isnumeric(Request("ActionType")) then
			p.Value = clng(Request("ActionType"))
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		set p =  cm.CreateParameter("@NewID", 3, &H0002)
		cm.Parameters.Append p
	            
		cm.Execute RowsEffected
				
		if cn.Errors.Count > 1 or RowsEffected <> 1 then
			cn.RollbackTrans
			strOut1 = ""
		else
			cn.CommitTrans
			strOut1 = cm("@NewID")
		end if
		
		set cm = nothing

	case "5"
		strOut1 = ""
		strOut2 = ""
		strOut3 = ""
		
		cn.BeginTrans
		
		set cm = server.CreateObject("ADODB.command")
		
		cm.ActiveConnection = cn
		cm.CommandText = "spUpdateExcelProfile"
		cm.CommandType =  &H0004
		
		set p =  cm.CreateParameter("@ID", 3, &H0001)
		p.value =  request("ID")
		cm.Parameters.Append p
	            
		Set p = cm.CreateParameter("@Products", 200, &H0001, 8000)
		p.Value = left(Request("Products"),8000)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Columns", 200, &H0001, 8000)
		p.Value = left(Request("Columns"),8000)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Header", 16, &H0001)
		if isnumeric(Request("Header")) then
			p.Value = clng(Request("Header"))
		else
			p.Value = 0
		end if
		cm.Parameters.Append p
	            
		cm.Execute RowsEffected
				
		if cn.Errors.Count > 1 or RowsEffected <> 1 then
			cn.RollbackTrans
			strOut1 = ""
		else
			cn.CommitTrans
			strOut1 = "1"
		end if
		
		set cm = nothing
		
		
	end select
	
	set rs = nothing
	cn.Close
	set cn = nothing



%>
<INPUT type="text" id=txtOut1 name=txtOut1 value="<%=strOut1%>">
<INPUT type="text" id=txtOut2 name=txtOut2 value="<%=strOut2%>">
<INPUT type="text" id=txtOut3 name=txtOut3 value="<%=strOut3%>">
</BODY>
</HTML>
