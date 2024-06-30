/***********************************************
* Pop-it menu- © Dynamic Drive (www.dynamicdrive.com)
* This notice MUST stay intact for legal use
* Visit http://www.dynamicdrive.com/ for full source code
***********************************************/

var defaultMenuWidth="150px" //set default menu width.
/*
Normally you would call a Javascript function from a link like this:
<a href="/nj.asp" onClick="cM_Cat(event, type, categoryid, groupid, palcategoryid);return false;" onmouseout="delayhidemenu();">My Link</a>
where nj.asp is a page telling the user that they need to have Javascript in order to work.

function cM_Cat(evt, type, CategoryID, GroupID, PALCategoryID) {
	var menu1=new Array()
	var wide = 150;
	menu1.push("<a href='/nj.asp' onClick=\"btnAddModules_onclick(" + type + "," + CategoryID + "," + GroupID + "," + PALCategoryID + "); return false;\">Add Modules</a>");
	kill();	// get rid of popup column hint. This is needed ONLY if you are using popup.inc which provides table cell popup hints.
	showmenu(evt, menu1.join(""), wide+'px');
}
*/
////No need to edit beyond here

var ie5=document.all && !window.opera
var ns6=document.getElementById

if (ie5||ns6)
	{	
		document.write('<div id="popitmenu" onMouseover="clearhidemenu();" onMouseout="dynamichide(event)"></div>')
	
	}


function iecompattest(){
	return (document.compatMode && document.compatMode.indexOf("CSS")!=-1)? document.documentElement : document.body
}

function showmenu(e, which, optWidth){
	if (!document.all&&!document.getElementById)
		return
	
	clearhidemenu()
	
	menuobj=ie5? document.all.popitmenu : document.getElementById("popitmenu")

	menuobj.innerHTML=which;	
	
	var iframeEl = document.createElement("IFRAME");
	iframeEl.frameBorder = 0;
	//iframeEl.src = "javascript:false;"; this will write a false in firefox, so comment out
	iframeEl.style.display = "none";
	iframeEl.style.position = "absolute";
	iframeEl.style.filter = "progid:DXImageTransform.Microsoft.Alpha(style=0,opacity=0)";
	menuobj.iframeEl = menuobj.parentNode.insertBefore(iframeEl, menuobj);

	menuobj.style.width=(typeof optWidth!="undefined")? optWidth : defaultMenuWidth
	menuobj.contentwidth=menuobj.offsetWidth
	menuobj.contentheight=menuobj.offsetHeight
	eventX=ie5? event.clientX : e.clientX
	eventY=ie5? event.clientY : e.clientY
	//Find out how close the mouse is to the corner of the window
	var rightedge=ie5? iecompattest().clientWidth-eventX : window.innerWidth-eventX
	var bottomedge=ie5? iecompattest().clientHeight-eventY : window.innerHeight-eventY

	//if the horizontal distance isn't enough to accomodate the width of the context menu
	if (rightedge<menuobj.contentwidth)
		//move the horizontal position of the menu to the left by it's width
		menuobj.style.left=ie5? iecompattest().scrollLeft+eventX-menuobj.contentwidth+"px" : window.pageXOffset+eventX-menuobj.contentwidth+"px"
	else
		//position the horizontal position of the menu where the mouse was clicked
		menuobj.style.left=ie5? iecompattest().scrollLeft+eventX+"px" : window.pageXOffset+eventX+"px"

	//same concept with the vertical position
	if (bottomedge<menuobj.contentheight)
		menuobj.style.top=ie5? iecompattest().scrollTop+eventY-menuobj.contentheight+"px" : window.pageYOffset+eventY-menuobj.contentheight+"px"
	else
		menuobj.style.top=ie5? iecompattest().scrollTop+event.clientY+"px" : window.pageYOffset+eventY+"px"
	
	if (menuobj.iframeEl != null)
	{   menuobj.iframeEl.style.left = menuobj.style.left;
		menuobj.iframeEl.style.top  = menuobj.style.top;
		menuobj.iframeEl.style.width  = menuobj.offsetWidth + "px";
		menuobj.iframeEl.style.height = menuobj.offsetHeight + "px";
		menuobj.iframeEl.style.display = "";
	}
	menuobj.style.visibility="visible"
	
	
	return false
}

function contains_ns6(a, b) {
	//Determines if 1 element in contained in another- by Brainjar.com
	while (b.parentNode)
		if ((b = b.parentNode) == a)
			return true;

	return false;
}

function hidemenu(){
	var menuobj=ie5? document.all.popitmenu : document.getElementById("popitmenu");
	if (menuobj)
	{	if (menuobj.iframeEl != null)
		{	
			menuobj.iframeEl.style.display = "none";
		}
		menuobj.style.visibility="hidden";
	}

}

function dynamichide(e){
	if (ie5&&!menuobj.contains(e.toElement))
		hidemenu()
	else if (ns6&&e.currentTarget!= e.relatedTarget&& !contains_ns6(e.currentTarget, e.relatedTarget))
		hidemenu()
}

function delayhidemenu(){
	delayhide=setTimeout("hidemenu()",500)
	
}

function clearhidemenu(){
	if (window.delayhide)
		clearTimeout(delayhide)
}

//if (ie5||ns6)
//	document.onclick=hidemenu
