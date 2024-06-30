<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<!-- #include file="../includes/bundleConfig.inc" -->
<script type="text/javascript" src="includes/client/json2.js"></script>
<script type="text/javascript" src="includes/client/json_parse.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

//-->
</SCRIPT>

<STYLE>
h3
    {
        font-family: Verdana;
        font-size:x-small;
    }
td,textarea,input,select
{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
  fieldset
  {
      background-color:White;
      padding-left:5px;
      padding-right:5px;
      padding-top:5px;
      padding-bottom:5px;
      margin: 10px 5px 10px 5px;
  }
a:link, a:visited
{
    color: blue;
}

a:hover
{
    color: red;
    text-decoration: none;
}
    
</STYLE>

<script language="javascript">


function window_onload() {
	//frames.myEditor.document.body.contentEditable = "True";
	//frames.myEditor.document.body.innerHTML = "<font face=verdana size=2>" + frmMain.txtArticleText.value + "</font>";
    //frames.myEditor.focus();
    modalDialog.load();
}



function disableEnterKey(e)
{
     var key;

     if(window.event)
          key = window.event.keyCode;     //IE
     else
          key = e.which;     //firefox

     if(key == 13)
          return false;
     else
          return true;
}

    function AddProject(){
        var strNewName = prompt("Enter the name of your Project","");
        var i;

        if (strNewName !=null && strNewName != "")
            {
            for (i=0;i<frmMain.cboProject.length;i++)
                if (frmMain.cboProject.options[i].text.toLowerCase() == strNewName.toLowerCase())
                    {
                    frmMain.cboProject.selectedIndex=i;
                    return;
                    }
            
            frmMain.cboProject.options[frmMain.cboProject.length] = new Option(strNewName, '0');
            frmMain.cboProject.selectedIndex = frmMain.cboProject.length-1;
            }

    }


    function cmdAddEmail_onclick() {

        //var url = "../Email/AddressBook.asp?AddressList=" + frmMain.txtNotify.value;
        var url = "/pulsarplus/core/User/AddressBook?AddressList=" + frmMain.txtNotify.value + "&PageName=TicketCategory"
        modalDialog.open({ dialogTitle: 'Address Book', dialogURL: '' + url + '', dialogHeight: 450, dialogWidth: 500, dialogResizable: true, dialogDraggable: true });
        $(".ui-dialog").css('display', 'block');
        $(".ui-widget-overlay").css('display', 'block');
        $("#modal_iframe").attr('src', url);
        globalVariable.save('txtNotify', 'email_field');
    }

    function cmdAddEmail_onclick_return(strResult) {
        if (typeof (strResult) != "undefined")
            frmMain.txtNotify.value = strResult;
    }

    function AddEmailCallBack(strResult) {
        if (typeof (strResult) != "undefined")
            frmMain.txtNotify.value = strResult;
    }
    function CloseDialog() {
        $(".ui-dialog").css('display', 'none');
        $(".ui-widget-overlay").css('display', 'none');
        $("#modal_iframe").attr('src', "");
    };

function ShowExample(ID){
    if (ID==1)
      frmMain.txtRequired.value="ERROR DISPLAYED:  \rPAGE NAME:   \rHOW TO REPRODUCE:   "
    else if (ID==2)
      frmMain.txtRequired.value="ERROR DISPLAYED:______\rPAGE NAME:______\rHOW TO REPRODUCE:______"
    else
        frmMain.txtRequired.value="Please provide a detailed explanation of how to reproduce this issue:"
}
</script>

</HEAD>


<BODY bgcolor="Ivory" onload="window_onload();">
<h3>Support Categories</h3>
<%

    
	dim cn, rs, cm
    dim blnFound
    dim strName
    dim strProjectID
    dim blnActive 
    dim strActiveChecked
    dim strOwnerID
    dim strOwnerName
    dim strRequired

    set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

    if request("ID") = "" then
        strName = ""
        strProjectID = ""
        blnActive = false
        strOwnerID = ""
        strOwnerName = ""
        strNotification = ""
        strRequired = ""
    else
        rs.open "spSupportCategorySelect " & clng(request("ID")),cn
        if not(rs.eof and rs.bof) then
            strName = rs("name") & ""
            strProjectID = rs("SupportProjectID") & ""
            blnActive = rs("Active") & ""
            strOwnerID = rs("OwnerID") & ""
            strOwnerName = rs("OwnerName") & ""
            strRequired = rs("RequiredFields") & ""
            strNotification = rs("NotificationList") & ""
        end if
        rs.Close
    end if
    
    if request("ID") = "" then
        blnActive=true
        strActiveChecked = " checked "
    elseif blnActive then
        strActiveChecked = " checked "
    else
        strActiveChecked = ""
    end if



    rs.open "spSupportProjectsListSelect",cn
    if trim(request("ID")) = "" then
        ProjectList = "<option></option>"
    end if
    blnFound=false
    do while not rs.EOF
        if trim(strProjectID) = trim(rs("ID")) then
            ProjectList = ProjectList & "<option selected value=""" & rs("ID") & """>" & rs("Name") & "</option>"
            blnFound = true
        else
            ProjectList = ProjectList & "<option value=""" & rs("ID") & """>" & rs("Name") & "</option>"
        end if
        rs.MoveNext
    loop
    rs.close
    if (not blnFound) and trim(strProjectName) <> "" then
        ProjectList = ProjectList & "<option selected value=""" & strProjectID & """>" & strProjectName & "</option>"
    end if

    OwnerList = "<option selected value=""" & strOwnerID & """>" & strOwnerName & "</option>"
    rs.open "spSupportAdminSelect",cn
    do while not rs.EOF
        if trim(strOwnerID) <> trim(rs("ID")) then
            OwnerList = OwnerList & "<option value=" & rs("ID") & ">" & rs("Name") & "</option>"
        end if
        rs.MoveNext
    loop
    rs.Close

        

%>
    <form id="frmMain" method="post" action="CategorySave.asp">
    <table cellpadding=2 border=1 cellspacing=0 bordercolor=tan bgcolor="cornsilk" style="border-width:1px;width:100%">
        <tr>
            <td valign=top><b>Name:</b>&nbsp;<font color=red>*</font></td>
            <td width="100%">
                <input id="txtName" name="txtName" style="width:100%" type="text"  onkeypress="return event.keyCode!=13" maxlength=120  value="<%= server.htmlencode(strName)%>"/>
        </tr>
        <tr>
            <td><b>Default&nbsp;Owner:</b>&nbsp;<font color=red>*</font></td>
            <td>
                <select id="cboOwner" name="cboOwner" style="width:100%">
                    <%=OwnerList%>
                </select>
            </td>
        </tr>
        <tr>
            <td><b>Project:</b>&nbsp;<font color=red>*</font></td>
            <td width="100%"><table cellspacing=0 cellpadding=0 width="100%"><tr><td width="100%">
                <select id="cboProject" name="cboProject" style="width:100%">
                    <%=ProjectList%>
                </select>
                </td>
                <td>&nbsp;
                <input id="cmdAddProject" type="button" value="Add" onclick="AddProject();" />
                </td>
                </tr></table>
            </td>
        </tr>
        <tr>
                <td valign=top><b>Notification&nbsp;List:</b>&nbsp;</td>
                <td>
                <table cellpadding=0 cellspacing=0 width="100%"><tr><td  width="100%">
                <textarea style="width:100%" id="txtNotify" name="txtNotify" rows=2><%=strNotification%></textarea>
                </td>
                <td valign=top>
                    &nbsp;
                    <input id="cmdAddEmail" type="button" value="Add" onclick="cmdAddEmail_onclick();" />
                    </td>
                </tr></table>
                </td>
        </tr>
        <tr>
            <td valign=top><b>Required&nbsp;Info&nbsp;Template:</b>&nbsp;<br><br>
            <a href="javascript:ShowExample(1);">Example 1</a><br>
            <a href="javascript:ShowExample(2);">Example 2</a><br>
            <a href="javascript:ShowExample(3);">Example 3</a><br>
            </td>
            <td colspan=3><textarea id="txtRequired" name="txtRequired" rows=10 style="width:100%"><%= server.htmlencode(strRequired)%></textarea></td>
        </tr>
        <tr>
            <td><b>Active:</b></td>
            <td>
                <input id="chkActive" <%=strActiveChecked%> name="chkActive" type="checkbox" value="1"/>&nbsp;This&nbsp;category&nbsp;is&nbsp;active&nbsp;
            </td>
        </tr>
        
    </table>
    <input id="txtID" name="txtID" type="hidden" value="<%=request("ID")%>"/>
    <input id="txtProjectName" name="txtProjectName" type="hidden" value=""/>
    </form>
<%






    set rs = nothing
    cn.Close
    set cn = nothing
%>
</BODY>
</HTML>




