<!--
   ######################################################
   # JAVASCRIPT POPUPS ROUTINE VERSION #7 07-Feb-2001   #        
   # Written by Mike McGrath [mike_mcgrath@lineone.net] # 
   # PC-Tested for Netscape 3.04, 4.61, 6.0, & IE5.5    #
   # Note: Popups may not cover all form field inputs.  #
   # PLEASE RETAIN THIS NOTICE WHEN COPYING MY SCRIPT.  #
   # THIS SCRIPT IS COPYRIGHT OF MIKE MCGRATH 1998-2001 #
   ######################################################
-->
<script type="text/javascript">
<!-- Original: Mike McGrath  (mike_mcgrath@lineone.net) -->
<!-- Web Site: http://website.lineone.net/~mike_mcgrath -->
<!--

var Xoffset=-60;        // modify these values to ...
var Yoffset= 20;        // change the popup position.
var popwidth=130;       // popup width
var bcolor="darkgray";  // popup border color
var fcolor="black";     // popup font color
var fface="arial";    // popup font face
var fsize="10px";	// popup font size in pixels
var dwidth=0;

// create content box
document.write("<DIV ID='pup'></DIV>");

// id browsers
var iex=(document.all);
var nav=(document.layers);
var old=(navigator.appName=="Netscape" && !document.layers && !document.getElementById);
var n_6=(window.sidebar);

// assign object
var skin;
if(nav) skin=document.pup;
if(iex) skin=pup.style;
if(n_6) skin=document.getElementById("pup").style;

// park modifier
var yyy=-1000;

// capture pointer
if(nav)document.captureEvents(Event.MOUSEMOVE);
if(n_6) document.addEventListener("mousemove",get_mouse,true);
if(nav||iex)document.onmousemove=get_mouse;

// set dynamic coords
function get_mouse(e){
	var x,y;

	if(nav || n_6) x=e.pageX;
	if(iex) x=event.x+document.body.scrollLeft; 

	if(nav || n_6) y=e.pageY;
	if(iex)	{
		y=event.y;
		if(navigator.appVersion.indexOf("MSIE 4")==-1)
		y+=document.body.scrollTop;
	}

	if(iex || nav)	{
		skin.top=y+yyy;
		skin.left=x+Xoffset; 
	}

	if(n_6)	{
		skin.top=(y+yyy)+"px";
		skin.left=x+Xoffset+"px";
	}    
	nudge(x);
}

// avoid edge overflow
function nudge(x) {
	var extreme,overflow,temp;

	// right
	if(iex) extreme=(document.body.clientWidth-popwidth);
	if(n_6 || nav) extreme=(window.innerWidth-popwidth);

	if(parseInt(skin.left)>extreme)	{
	
		overflow=parseInt(skin.left) - extreme;
		if( parseInt(skin.left) >= ( documentwidth() ) ) { temp = documentwidth(); }
		else if ( document.body.scrollLeft < 1 ) temp=parseInt(skin.left) - overflow;
		else temp=parseInt(skin.left);
		if(nav || iex)skin.left=temp;
		if(n_6)skin.left=temp+"px";
	}

	// left
	if(parseInt(skin.left)<1) {
		overflow=parseInt(skin.left)-1;
		temp=parseInt(skin.left);
		temp-=overflow;
		if(nav || iex) skin.left=temp;
		if(n_6)skin.left=temp+"px";
	}
}

function documentwidth()
{
	if( dwidth == 0 )
		dwidth = document.body.scrollWidth - 130;
	else if( document.body.scrollWidth < dwidth )
		dwidth = document.body.scrollWidth - 130;
		
	return dwidth;
}

// write content & display
function popup(msg,w,bak) {
	popwidth = w;

	var content="<TABLE WIDTH='"+popwidth+"' BORDER='1' BORDERCOLOR="+bcolor+" CELLPADDING=2 CELLSPACING=0 "+"BGCOLOR="+bak+">"
	content += "<TD ALIGN='center'>"
	content += msg + "</TD></TABLE>";

	if(old) {
		alert(msg);
		return;
	} 
   
	yyy=Yoffset; 
	skin.width=popwidth;
	alert

	if(nav)	{ 
		skin.document.open();
		skin.document.write(content);
		skin.document.close();
		skin.visibility="visible";
	}

	if(iex)	{        
		pup.innerHTML=content;
		skin.visibility="visible";
	}  

	if(n_6)	{   
		document.getElementById("pup").innerHTML=content;
		skin.visibility="visible";
	}
}


// park content box
function kill(){
	if(!old) {
		yyy=-1000;
		skin.visibility="hidden";
		skin.width=0;
	}
}

function set_Yoffset(lng)
{

	Yoffset = lng;


}

//-->
</script>