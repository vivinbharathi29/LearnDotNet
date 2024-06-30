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
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
    var OutArray = new Array;
    if (typeof (txtSuccess) != "undefined")
		{
		    if (txtSuccess.value != "") {
		        OutArray = txtSuccess.value.split("||")

		        window.parent.returnValue = OutArray;
		        window.parent.close();
		    }
		    else {
		        alert("Unable to find the selected deliverable.");
		        window.parent.close();
		    }
        }

}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">


<%
    dim cn, rs
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
    
    dim strID

    if request("optType") = "1" then
        strID = request("txtID")
    else
        if request("txtAllVersions") <> "" then
            strID = request("txtAllVersions")
        end if
        if request("chkVersion") <> "" then
            if strID = "" then
                strId = request("chkVersion")
            else
                strId = strId & ", " & request("chkVersion")
            end if
        end if
    end if



    dim TestArray

    if trim(strID) = "" then
        isValid = false
        response.write "No versions selected."
    else
        TestArray = split(strID,",")
        for i = 0 to ubound(testArray)
            if trim(testarray(i)) = "" or (not isnumeric(trim(testArray(i)))) then
                respnose.write "Invalid ID numbers detected."
                strID = ""
                exit for
            end if
        next
    end if

    
    'Lookup deliverable info here and format output.
    rs.open "Select ID, deliverableName as name, Version, Revision, Pass from deliverableversion where id in (" & strID & ") order by name, id",cn
    strSuccess = ""
    do while not rs.eof
        strversion = rs("Version") & ""
        if trim(rs("revision") & "") <> "" then
            strversion = strversion & "," & rs("revision")
        end if
        if trim(rs("pass") & "") <> "" then
            strversion = strversion & "," & rs("pass")
        end if
        strSuccess = strSuccess & rs("ID") & "^^" & rs("Name") & " [" & strversion & "]||"
        rs.movenext
    loop
    rs.close



    set rs = nothing
    cn.close
    set cn = nothing


%>

<textarea id="txtSuccess" name=txtSuccess rows="5" style="display:none;width:100%"><%=strSuccess%></textarea>
</BODY>
</HTML>

