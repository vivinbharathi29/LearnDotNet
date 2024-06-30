<%@ Language=VBScript %>
<!DOCTYPE html>
<HTML>
<HEAD>
<% dim AppRoot : AppRoot = Session("ApplicationRoot") %>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="<%= AppRoot %>/style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
<script src="<%= AppRoot %>/includes/client/jquery.min.js" type="text/javascript"></script>
<script src="<%= AppRoot %>/includes/client/jquery-ui.min.js" type="text/javascript"></script>
 <script src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var KeyString = "";

function combo_onkeypress() {
    //detect if key press is enter button from keyboard, if so trigger function that is called when ok button is clicked.
	if (event.keyCode == 13)
	{
		KeyString = "";
		frmEmployee.submit();
		window.top.location.href = window.top.location.href;
	}
	else if (event.keyCode == 8)
	{
	    return false;
	}

	//else
	//	{
	//	KeyString=KeyString+ String.fromCharCode(event.keyCode);
	//	event.keyCode = 0;
	//	var i;
	//	var regularexpression;
		
	//	for (i=0;i<event.srcElement.length;i++)
	//		{
	//			regularexpression = new RegExp("^" + KeyString,"i")
	//			if (regularexpression.exec(event.srcElement.options[i].text)!=null)
	//				{
	//				event.srcElement.selectedIndex = i;
	//				};
				
	//		}
	//	return false;
	//	}	
}

function combo_onfocus() {
	KeyString = "";
}

function combo_onclick() {
	KeyString = "";
}

function combo_onkeydown() {
    //when window is called from jquery dialogue, backspace key will not be detect from onkeypresss so we need this
	if (event.keyCode==8)
	{
		//if (String(KeyString).length >0)
		//	KeyString= Left(KeyString,String(KeyString).length-1);
		return false;
	}
}

function Left(str, n)
    {
	if (n <= 0)     // Invalid bound, return blank string
		return "";
    else if (n > String(str).length)   // Invalid bound, return
        return str;                // entire string
    else // Valid bound, return appropriate substring
        return String(str).substring(0,n);
    }


function cmdCancel_onclick() {
    CloseIframeDialog();
}

function cmdReset_onclick() {
	frmEmployee.cboEmployee.selectedIndex=0;
	frmEmployee.submit();
	window.top.location.href = window.top.location.href;
}

function cmdOK_onclick() {
	frmEmployee.submit();
    window.top.location.href = window.top.location.href;
}

function window_onload() {
	frmEmployee.cboEmployee.focus();
}

function CloseIframeDialog() {
    var iframeName = window.name;
    if (iframeName != '') {
        if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        }
        else {
            parent.window.parent.ClosePropertiesDialog();
        }
        //parent.window.parent.ClosePropertiesDialog();
    } else {
        if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        }
        else {
            window.close();
        }
        //window.close();
    }
}

//-->
</SCRIPT>
      <style>
  .custom-combobox {
    position: relative;
    display: inline-block;
  }
  .custom-combobox-toggle {
    position: absolute;
    top: 0;
    bottom: 0;
    margin-left: -1px;
    padding: 0;
  }
  .custom-combobox-input {
    margin: 0;
    padding: 5px 10px;
  }
  </style>
  <script>
  (function( $ ) {
    $.widget( "custom.combobox", {
      _create: function() {
        this.wrapper = $( "<span>" )
          .addClass( "custom-combobox" )
          .insertAfter( this.element );
 
        this.element.hide();
        this._createAutocomplete();
        this._createShowAllButton();
      },
 
      _createAutocomplete: function() {
        var selected = this.element.children( ":selected" ),
          value = selected.val() ? selected.text() : "";
 
        this.input = $( "<input>" )
          .appendTo( this.wrapper )
          .val( value )
          .attr( "title", "" )
          .addClass( "custom-combobox-input ui-widget ui-widget-content ui-state-default ui-corner-left" )
          .autocomplete({
            delay: 0,
            minLength: 0,
            source: $.proxy( this, "_source" )
          })
          .tooltip({
            tooltipClass: "ui-state-highlight"
          });
 
        this._on( this.input, {
          autocompleteselect: function( event, ui ) {
            ui.item.option.selected = true;
            this._trigger( "select", event, {
              item: ui.item.option
            });
          },
 
          autocompletechange: "_removeIfInvalid"
        });
      },
 
      _createShowAllButton: function() {
        var input = this.input,
          wasOpen = false;
 
        $( "<a>" )
          .attr( "tabIndex", -1 )
          .attr( "title", "Show All Items" )
          .tooltip()
          .appendTo( this.wrapper )
          .button({
            icons: {
              primary: "ui-icon-triangle-1-s"
            },
            text: false
          })
          .removeClass( "ui-corner-all" )
          .addClass( "custom-combobox-toggle ui-corner-right" )
          .mousedown(function() {
            wasOpen = input.autocomplete( "widget" ).is( ":visible" );
          })
          .click(function() {
            input.focus();
 
            // Close if already visible
            if ( wasOpen ) {
              return;
            }
 
            // Pass empty string as value to search for, displaying all results
            input.autocomplete( "search", "" );
          });
      },
 
      _source: function( request, response ) {
        var matcher = new RegExp( $.ui.autocomplete.escapeRegex(request.term), "i" );
        response( this.element.children( "option" ).map(function() {
          var text = $( this ).text();
          if ( this.value && ( !request.term || matcher.test(text) ) )
            return {
              label: text,
              value: text,
              option: this
            };
        }) );
      },
 
      _removeIfInvalid: function( event, ui ) {
 
        // Selected an item, nothing to do
        if ( ui.item ) {
          return;
        }
 
        // Search for a match (case-insensitive)
        var value = this.input.val(),
          valueLowerCase = value.toLowerCase(),
          valid = false;
        this.element.children( "option" ).each(function() {
          if ( $( this ).text().toLowerCase() === valueLowerCase ) {
            this.selected = valid = true;
            return false;
          }
        });
 
        // Found a match, nothing to do
        if ( valid ) {
          return;
        }
 
        // Remove invalid value
        this.input
          .val( "" )
          .attr( "title", value + " didn't match any item" )
          .tooltip( "open" );
        this.element.val( "" );
        this._delay(function() {
          this.input.tooltip( "close" ).attr( "title", "" );
        }, 2500 );
        this.input.autocomplete( "instance" ).term = "";
      },
 
      _destroy: function() {
        this.wrapper.remove();
        this.element.show();
      }
    });
  })( jQuery );
 
  $(function() {
    $( "#combobox" ).combobox();
    $( "#toggle" ).click(function() {
      $( "#combobox" ).toggle();
    });
  });
  </script>
</HEAD>
<STYLE>
td{
	FONT-FAMILY=Verdana;
	FONT-SIZE=x-small;
}
</STYLE>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">

<%
		dim CurrentDomain
		dim CurrentUser
		dim strEmployees
		dim ActualUserID
        dim rs
        dim blnActualUserPulsarAdmin

	    CurrentUser = lcase(Session("LoggedInUser"))
    
	    set rs = server.CreateObject("ADODB.recordset")

        if instr(currentuser,"\") > 0 then
		    CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		    Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	    end if

	    set cn = server.CreateObject("ADODB.Connection")
	    cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	    cn.Open

        set cm = server.CreateObject("ADODB.Command")
	    Set cm.ActiveConnection = cn
	    cm.CommandType = 4
	    cm.CommandText = "spGetEmployeeImpersonateID"
		
	    Set p = cm.CreateParameter("@NTName", 200, &H0001, 80)
	    p.Value = Currentuser
	    cm.Parameters.Append p
	
	    Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	    p.Value = CurrentDomain
	    cm.Parameters.Append p
	
	    rs.CursorType = adOpenForwardOnly
	    rs.LockType=AdLockReadOnly
	    Set	rs = cm.Execute 
	
	    set cm=nothing	
		
	    if not (rs.EOF and rs.BOF) then
            ActualUserID = rs("EmployeeID")
        else
            ActualUserID = ""
        end if
        rs.close

        dim blnSupportAdmin
        rs.open "spSupportIsAdminSelect " & clng(CurrentUserID),cn
	    'Response.write("UserName: " + rs("Name"))
        if rs.eof and rs.bof then
            blnSupportAdmin = false
        else
            blnSupportAdmin = true
        end if
        rs.close

	    set cm = server.CreateObject("ADODB.Command")
	    Set cm.ActiveConnection = cn
	    cm.CommandType = 4
	    cm.CommandText = "spGetUserInfo"
		
	
	    Set p = cm.CreateParameter("@UserName", 200, &H0001, 30)
	    p.Value = Currentuser
	    cm.Parameters.Append p
	
	    Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	    p.Value = CurrentDomain
	    cm.Parameters.Append p
	
	    rs.CursorType = adOpenForwardOnly
	    rs.LockType=AdLockReadOnly
	    Set	rs = cm.Execute 
	
	    set cm=nothing	
		
	    if not (rs.EOF and rs.BOF) then

            if CLng(rs("ID")) <> CLng(ActualUserID) or Cint(rs("PulsarSystemAdmin")) = 1 or blnSupportAdmin then
                blnActualUserPulsarAdmin = True
            else
                blnActualUserPulsarAdmin = False
            end if
        else
            blnActualUserPulsarAdmin = false
        end if
        rs.close

    rs.open "spSupportIsAdminSelect " & clng(CurrentUserID),cn
	'Response.write("UserName: " + rs("Name"))
    if rs.eof and rs.bof then
        blnSupportAdmin = false
    else
        blnSupportAdmin = true
    end if
    rs.close

    if blnActualUserPulsarAdmin = true or blnSupportAdmin = true then
	
%>
<form ID=frmEmployee action="ChangeUserSave.asp" method=post>
<TABLE width=100% bgcolor=cornsilk border=1 bordercolor=tan cellspacing=0 cellpadding=1>
<TR>
	<TD><b>Impersonate:</b>&nbsp;&nbsp;</TD>
	<TD width=100%>
        <SELECT style="width:100%" id="cboEmployee" name="cboEmployee" onkeypress="return combo_onkeypress()" onkeydown="return combo_onkeydown()">
			<OPTION value="0"></OPTION>
	<%
	dim strImpersonateID
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	set rs = server.CreateObject("ADODB.recordset")
	
	rs.Open "spGetEmployeeImpersonateID '" & CurrentUser & "','" & CurrentDomain & "'",cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		strImpersonateID = ""
		strEmployeeID=""
	else
		strImpersonateID = trim(rs("ImpersonateID") & "")
		strEmployeeID = rs("EmployeeID")
	end if
	rs.Close	
	rs.Open "spGetEmployees 1",cn,adOpenForwardOnly
	strEmployees = ""
	do while not rs.EOF
		if not(lcase(rs("Domain")) = CurrentDomain and lcase(rs("NTName")) = CurrentUser) then
            '05/10/16 Malichi - PBI 10931: User Admin; Update the Impersonate menu to Include email address in () when the names are the same
			if trim(rs("ID")) = strImpersonateID then
				Response.Write "<OPTION selected value=""" & rs("ID") & """>" & rs("Name") & rs("Email") & "</OPTION>" & vbcrlf
			else
				Response.Write "<OPTION value=""" & rs("ID") & """>" & rs("Name") & rs("Email") & "</OPTION>" & vbcrlf
			end if
		end if
		rs.MoveNext
	loop
	
	set rs = nothing	
	cn.Close
	set cn=nothing

%>
</select>
	</TD>
</TR>
</TABLE>

<TABLE width=100%><TR><TD align=right><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">&nbsp;
    <INPUT type="button" value="Reset" id=cmdReset name=cmdReset LANGUAGE=javascript onclick="return cmdReset_onclick()">&nbsp;
    <INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></td></tr></table>
<INPUT type="hidden" id=txtEmployeeID name=txtEmployeeID value="<%=strEmployeeID%>">
</form>

    <% 
        else
            Response.Write "You do not have access to impersonate a user."
        end if %>
</BODY>
</HTML>
