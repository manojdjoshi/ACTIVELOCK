/******************************************
* Popup Box- By Jim Silver @ jimsilver47@yahoo.com
* Visit http://www.dynamicdrive.com/ for full source code
* This notice must stay intact for use
******************************************
* modified by FRIEND SOFTWARE  (2006)
******************************************/

var ns4=document.layers;
var ie4=document.all;
var ns6=document.getElementById&&!document.all;
var opera8=DetectOpera();
var calunit=ns4? "" : "px";
var modulePath;
//drag drop function for NS 4////
/////////////////////////////////

var dragswitch=0
var nsx
var nsy
var nstemp

function DetectOpera()
{
	if(navigator.userAgent.indexOf("Opera")!=-1)
	{
		var versionindex=navigator.userAgent.indexOf("Opera")+6
		if (parseInt(navigator.userAgent.charAt(versionindex))>=8)
			return true;
	}
	return false;
}

function inimodulePath(mvalue)
{
	modulePath = mvalue;
}

function drag_dropns(name){
if (!ns4)
return
temp=eval(name)
temp.captureEvents(Event.MOUSEDOWN | Event.MOUSEUP)
temp.onmousedown=gons
temp.onmousemove=dragns
temp.onmouseup=stopns
}

function gons(e){
temp.captureEvents(Event.MOUSEMOVE)
nsx=e.x
nsy=e.y
}
function dragns(e){
if (dragswitch==1){
temp.moveBy(e.x-nsx,e.y-nsy)
return false
}
}

function stopns(){
temp.releaseEvents(Event.MOUSEMOVE)
}

//drag drop function for ie4+ and NS6////
/////////////////////////////////


function drag_drop(e){
if (ie4&&dragapproved){
crossobj.style.left=tempx+event.clientX-offsetx
crossobj.style.top=tempy+event.clientY-offsety
return false
}
else if (ns6&&dragapproved){
crossobj.style.left=tempx+e.clientX-offsetx+"px"
crossobj.style.top=tempy+e.clientY-offsety+"px"
return false
}
}

function initializedrag(e){
crossobj=ns6? document.getElementById("showimage") : document.all.showimage
var firedobj=ns6? e.target : event.srcElement
var topelement=ns6? "html" : document.compatMode && document.compatMode!="BackCompat"? "documentElement" : "body"
while (firedobj.tagName!=topelement.toUpperCase() && firedobj.id!="dragbar"){
firedobj=ns6? firedobj.parentNode : firedobj.parentElement
}

if (firedobj.id=="dragbar"){
offsetx=ie4? event.clientX : e.clientX
offsety=ie4? event.clientY : e.clientY

tempx=parseInt(crossobj.style.left)
tempy=parseInt(crossobj.style.top)

dragapproved=true
document.onmousemove=drag_drop
}
}
document.onmouseup=new Function("dragapproved=false")

////drag drop functions end here//////

function hidebox(){
crossobj=ns6? document.getElementById("showimage") : document.all.showimage
if (ie4||ns6)
crossobj.style.visibility="hidden"
else if (ns4)
document.showimage.visibility="hide"
}

function showbox(){
crossobj=ns6? document.getElementById("showimage") : document.all.showimage
if (ie4||ns6)
crossobj.style.visibility="visible"
else if (ns4)
document.showimage.visibility="visible"
}

//create popup from here
function createPopup(ctx, msgTitle, msgStr, mWidth, mHeight)
{
	var sTableID = 'test';
	var mypopup;
	if (document.getElementById("showimage")!=null)
	{
		var mypopup = document.getElementById("showimage");
	}
	else{
		var mypopup = document.createElement('DIV');
		document.body.appendChild(mypopup);
	}

	mypopup.setAttribute('name', 'showimage');
	mypopup.setAttribute('id', 'showimage');
	mypopup.style.position="absolute";
	mypopup.style.zIndex=50;
	mypopup.style.width=mWidth+"px";
	mypopup.style.height=mHeight+"px";

	if (ie4){documentWidth  =truebody().offsetWidth/2+truebody().scrollLeft-20;
	documentHeight =truebody().offsetHeight/2+truebody().scrollTop-20;}	
	else if (ns4){documentWidth=window.innerWidth/2+window.pageXOffset-20;
	documentHeight=window.innerHeight/2+window.pageYOffset-20;} 
	else if (ns6){documentWidth=self.innerWidth/2+window.pageXOffset-20;
	documentHeight=self.innerHeight/2+window.pageYOffset-20;} 
	
	mypopup.style.top = documentHeight - mHeight/2 + calunit;
	mypopup.style.left = documentWidth - mWidth/2  + calunit;
	mypopup.style.border="solid 1px black";


	mypopup.innerHTML = "<table border='0' cellspacing='0' cellpadding='0' width='100%' height='100%' class='popupdivtable'><tr><td id='dragbar' class='popupdivtitle' width='100%' onMousedown='initializedrag(event)'>" + 
			"<ilayer width='100%' onSelectStart='return false'><layer width='100%' onMouseover='dragswitch=1;if (ns4) drag_dropns(showimage)' onMouseout='dragswitch=0' id='msgTitle'>" + msgTitle + "</layer></ilayer></td><td class='popupdivtitle' style='width:16px;' ><a href='#' onClick='hidebox();return false'><img src='" + modulePath + "close.gif' width='16px' height='14px' border=0></a></td></tr>" + 
			"<tr><td class='popupdivcontent' colspan=2 id='msgStr'>" + msgStr + "</td><td></td></tr></table>";


	if (ie4 && !opera8)
	{ 
		mypopup.insertAdjacentHTML("afterBegin", '<IFRAME style="position: absolute;z-index:-1;border: none;margin: 0px 0px 0px 0px; background-color:Transparent;" src="javascript:false;" frameBorder="0" scrolling="no" id="' + sTableID + '_hvrShm" />'); 
		var iframeShim = document.getElementById(sTableID + "_hvrShm"); 
		iframeShim.style.top = -1; 
		iframeShim.style.left = -1; 
		iframeShim.style.width = mypopup.style.width;
		iframeShim.style.height = mypopup.style.height;
	} 

	showbox();
}

// Get object left position, even if nested
function getAbsLeft(o) {
	oLeft = o.offsetLeft
	while(o.offsetParent!=null) {
		oParent = o.offsetParent
		oLeft += oParent.offsetLeft
		o = oParent
	}
	return oLeft
}

// Get object top position, even if nested
function getAbsTop(o) {
	oTop = o.offsetTop
	while(o.offsetParent!=null) {
		oParent = o.offsetParent
		oTop += oParent.offsetTop
		o = oParent
	}
	return oTop
}

function truebody(){
return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body
}
