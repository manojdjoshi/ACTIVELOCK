// Desarrollado por: Ricaute Jiménez Sánchez (La Chorrera, Panamá)
// Componente: PopCalendar.js(3.2.8)
// Correo: ricaj0625@yahoo.com
// Última actualización: 10 de mayo de 2006
// Por favor mantenga estos créditos (Please, keep these credits)

if(typeof(objPopCalList)=='undefined')
{
	var objPopCalList=[]
	var PopCalendar={majorVersion:3,minorVersion:2.8,newCalendar:new Function("_name","return(PopCalGetCalendarInstance(_name))")}
}

function PopCalGetCalendarInstance(_name)
{
	var _obj=new PoPCalCreateCalendarInstance(_name)
	objPopCalList[_obj.id]=_obj
	return(_obj)
}

function PoPCalCreateCalendarInstance(_name)
{
	var _id=PopCalGetCalendarIndex(_name)
	this.id=_id
	this.calendarName=_name
	this.BlankFieldText=""
	this.ClientScriptOnDateChanged=""
	this.ClientScriptDisabledDateStyle=""
	this.ClientScriptHolidayDateStyle=""
	this.ClientScriptWeekendDateStyle=""
	this.ClientScriptSpecialDateStyle=""
	this.startAt=0
	this.weekend="60"
	this.showWeekNumber=0
	this.weekNumberFormula=0
	this.showDaysOutMonth=0
	this.showToday=1
	this.showWeekend=0
	this.showHolidays=1
	this.showSpecialDay=1
	this.selectWeekend=1
	this.selectHoliday=1
	this.addCarnival=0
	this.addGoodFriday=0
	this.buttons=0
	this.clientValidator=1
	this.defaultFormat="dd-mm-yyyy"
	this.fixedX=-1
	this.fixedY=-1
	this.fade=0
	this.shadow=0
	this.move=0
	this.saveMovePos=0
	this.centuryLimit=40
	this.keepInside=1
	this.executeFade=true
	this.forceTodayTo=null
	this.forceTodayFormat=null
	this.imgDir=""
	this.CssClass=""
	this.todayString=""
	this.lr=1
	this.CarnivalString="Carnival"
	this.GoodFridayString="Good Friday"
	this.selectDateMessage=""
	this.monthSelected=null
	this.yearSelected=null
	this.dateSelected=null
	this.omonthSelected=null
	this.oyearSelected=null
	this.odateSelected=null
	this.monthConstructed=null
	this.yearConstructed=null
	this.intervalID1=null
	this.intervalID2=null
	this.timeoutID1=null
	this.timeoutID2=null
	this.timeoutID3=null
	this.ctlId="n:1"
	this.ctlIdNow="n:2"
	this.ctl=new Function("return(PopCalGetById(objPopCalList["+_id+"].ctlId))")
	this.dateFormat=null
	this.nStartingYear=null
	this.onKeyPress=null
	this.onClick=null
	this.onSelectStart=null
	this.onContextMenu=null
	this.onmousemove=null
	this.onmouseup=null
	this.onresize=null
	this.onscroll=null
	this.ControlAlignLeft=null
	this.ie=false
	this.ieVersion=0
	this.dom=document.getElementById
	this.ns4=document.layers
	this.opera=navigator.userAgent.indexOf("Opera")!=-1
	this.mozilla=((navigator.userAgent.indexOf("Mozilla")!=-1)&&(navigator.userAgent.indexOf("Netscape")==-1))
	if(!this.opera)
	{
		this.ie=document.all
		var ms=navigator.appVersion.indexOf("MSIE")
		if(ms!=-1) this.ieVersion=parseFloat(navigator.appVersion.substring(ms+5,ms+8))
	}
	this.dateFrom=01
	this.monthFrom=00
	this.yearFrom=1900
	this.dateUpTo=31
	this.monthUpTo=11
	this.yearUpTo=2099
	this.oDate=null
	this.oMonth=null
	this.oYear=null
	this.countMonths=12
	this.today=null
	this.dayNow=0
	this.dateNow=0
	this.monthNow=0
	this.yearNow=0
	this.defaultX=0
	this.defaultY=0
	this.keepMonth=false
	this.keepYear=false
	this.bShow=false
	this.PopCalTimeOut=null
	this.PopCalDragClose=false
	this.HalfYearList=5
	this.HolidaysCounter=0
	this.Holidays=[]
	this.movePopCal=false
	this.commandExecute=null
	this.Object={initialized:0}
	this.initCalendar=new Function("PopCalInitCalendar("+_id+");")
	this.show=new Function("ctl","format","from","to","execute","PopCalShow(ctl,format,from,to,execute,"+_id+");")
	this.addHoliday=new Function("d","m","y","t","PopCalAddHoliday(d,m,y,t,"+_id+");")
	this.addIrregularHoliday=new Function("s","dw","m","t","PopCalAddIrregularHoliday(s,dw,m,t,"+_id+");")
	this.addSpecialDay=new Function("d","m","y","t","PopCalAddSpecialDay(d,m,y,t,"+_id+");")
	this.addIrregularSpecialDay=new Function("s","dw","m","t","PopCalAddIrregularSpecialDay(s,dw,m,t,"+_id+");")
	this.addRecurrenceSpecialDay=new Function("d","m","y","i","f","r","t","PopCalAddRecurrenceSpecialDay(d,m,y,i,f,r,t,"+_id+");")
	this.formatDate=new Function("dateValue","oldFormat","newFormat","return(PopCalFormatDate(dateValue,oldFormat,newFormat,"+_id+"));")
	this.addDays=new Function("dateValue","format","daysToAdd","return(PopCalAddDays(dateValue,format,daysToAdd,"+_id+"));")
	this.forcedToday=new Function("dateValue","format","PopCalForcedToday(dateValue,format,"+_id+");")
	this.getDate=new Function("dateValue","dateFormat","return(PopCalGetDate(dateValue,dateFormat,"+_id+"));")
	this.getWeekNumber=new Function("dateValue","return(PopCalGetWeekNumber1(dateValue,"+_id+"));")
	this.scroll=new Function("PopCalScroll("+_id+");")
	this.hide=new Function("PopCalHideCalendar("+_id+",true)")
	this.isGoodFriday=new Function("dateValue","return(PopCalIsGoodFriday(dateValue));")
	this.isCarnival=new Function("dateValue","return(PopCalIsCarnival(dateValue));")
}

function PopCalGetCalendarIndex(_name)
{
	for(var i=0;i<objPopCalList.length;i++)
	{
		if(objPopCalList[i].calendarName==_name)
		{
			return(i)
		}
	}
	return(objPopCalList.length)
}

function PopCalInitCalendar(l)
{
	var objPopCal=objPopCalList[l]
	var PopCal=objPopCal.Object
	if(PopCal)
	{
		var _FirstNode=document.body.childNodes[0]
		var _Component=null
		var _Style=null
		if(PopCal.initialized==0)
		{
			if(!PopCalGetById("popupSuperYear"))
			{
				if(objPopCal.ie)
				{
					var sComponents="<div id='CalendarLoadFilters' style='z-index:+100000;position:absolute;top:0px;left:0px;display:none;filter="+'"'+"alpha() blendTrans()"+'"'+"'></div>"
					if(objPopCal.ieVersion>=5.5)
					{
						sComponents+="<iframe id='popupOverShadow' src='javascript:false;' scrolling=no frameborder=0 style='position:absolute;left:0px;top:0px;width:0px;height:0px;z-index:+10000;display:none;filter:progid:DXImageTransform.Microsoft.Alpha(opacity=0);'></iframe>"
						sComponents+="<iframe id='popupOverCalendar' src='javascript:false;' scrolling=no frameborder=0 style='position:absolute;left:0px;top:0px;width:0px;height:0px;z-index:+10000;display:none;filter:progid:DXImageTransform.Microsoft.Alpha(opacity=0);'></iframe>"
						sComponents+="<iframe id='popupOverYearMonth' src='javascript:false;' scrolling=no frameborder=0 style='position:absolute;left:0px;top:0px;width:0px;height:0px;z-index:+10000;display:none;filter:progid:DXImageTransform.Microsoft.Alpha(opacity=0);'></iframe>"
					}
					document.body.insertAdjacentHTML("afterBegin",sComponents)
				}
				_Component=document.createElement("DIV")
				_Style=_Component.style
				_Component.id="popupSuperShadowRight"
				_Component.onclick=new Function('PopCalCalendarVisible().bShow=true;')
				_Style.zIndex=+1000000
				_Style.position="absolute"
				_Style.top="0px"
				_Style.left="0px"
				_Style.fontSize="10px"
				_Style.width="10px"
				_Style.visibility="hidden"
				_Style.backgroundColor="black"
				document.body.insertBefore(_Component, _FirstNode)
				_Component=document.createElement("DIV")
				_Style=_Component.style
				_Component.id="popupSuperShadowBottom"
				_Component.onclick=new Function('PopCalCalendarVisible().bShow=true;')
				_Style.zIndex=+1000000
				_Style.position="absolute"
				_Style.top="0px"
				_Style.left="0px"
				_Style.fontSize="10px"
				_Style.height="10px"
				_Style.visibility="hidden"
				_Style.backgroundColor="black"
				document.body.insertBefore(_Component, _FirstNode)
				_Component=document.createElement("DIV")
				_Style=_Component.style
				_Component.id="popupSuperMonth"
				_Component.onclick=new Function('PopCalCalendarVisible().bShow=true;')
				_Style.zIndex=+1000000
				_Style.position="absolute"
				_Style.top="0px"
				_Style.left="0px"
				_Style.display="none"
				document.body.insertBefore(_Component, _FirstNode)
				_Component=document.createElement("DIV")
				_Style=_Component.style
				_Component.id="popupSuperYear"
				_Component.onclick=new Function('PopCalCalendarVisible().bShow=true;')
				_Component.onmousewheel=new Function('PopCalWheelYear(PopCalCalendarVisible().id)')
				_Style.zIndex=+1000000
				_Style.position="absolute"
				_Style.top="0px"
				_Style.left="0px"
				_Style.display="none"
				document.body.insertBefore(_Component, _FirstNode)
			}
			var _id=objPopCal.id
			PopCal.id=_id
			if((objPopCal.centuryLimit<0)||(objPopCal.centuryLimit>99))
			{
				objPopCal.centuryLimit=40
			}
			objPopCal.today=new Date()
			if(objPopCal.forceTodayTo)
			{
				if(objPopCal.forceTodayFormat==null)
				{
					objPopCal.forceTodayFormat=objPopCal.defaultFormat
				}
				if(PopCalSetDMY(objPopCal.forceTodayTo,objPopCal.forceTodayFormat,l))
				{
					objPopCal.today=new Date(objPopCal.oYear,objPopCal.oMonth,objPopCal.oDate)
				}
			}
			objPopCal.dayNow=objPopCal.today.getDay()
			objPopCal.dateNow=objPopCal.today.getDate()
			objPopCal.monthNow=objPopCal.today.getMonth()
			objPopCal.yearNow=objPopCal.today.getFullYear()
			objPopCal.monthConstructed=false
			objPopCal.yearConstructed=false
			var _leftButton=''
			var _rightButton=''
			_Component=document.createElement("DIV")
			_Style=_Component.style
			_Component.id="popupSuperCalendar"+_id
			_Component.className=objPopCal.CssClass
			_Component.oncontextmenu=new Function('return(false);')
			_Component.onclick=new Function('PopCalDownMonth('+l+');PopCalDownYear('+l+');objPopCalList['+l+'].bShow=true;')
			_Style.zIndex=+100000
			_Style.position="absolute"
			_Style.top="0px"
			_Style.left="0px"
			_Style.visibility="hidden"
			document.body.insertBefore(_Component, _FirstNode)
			if(objPopCal.lr==1)
			{
				_leftButton="<span id='popupSuperSpanLeft"+_id+"' class='DropDownStyle' ondrag='return(false)' onmouseover='PopCalSwapImage(\"popupSuperChangeLeft"+_id+"\",\"left2.gif\","+l+");this.className=\"DropDownOverStyle\";' onmouseout='clearInterval(objPopCalList["+l+"].intervalID1);PopCalSwapImage(\"popupSuperChangeLeft"+_id+"\",\"left1.gif\","+l+");this.className=\"DropDownOutStyle\";window.status=\"\"' onclick='PopCalDecMonth("+l+")' onmousedown='clearTimeout(objPopCalList["+l+"].timeoutID1);objPopCalList["+l+"].timeoutID1=setTimeout(\"PopCalStartDecMonth("+l+")\",100)' onmouseup='clearTimeout(objPopCalList["+l+"].timeoutID1);clearInterval(objPopCalList["+l+"].intervalID1)'>&nbsp;<img id='popupSuperChangeLeft"+_id+"' src='"+objPopCal.imgDir+"left1.gif' border='0' />&nbsp;</span>"
				_rightButton="<span id='popupSuperSpanRight"+_id+"' class='DropDownStyle' ondrag='return(false)' onmouseover='PopCalSwapImage(\"popupSuperChangeRight"+_id+"\",\"right2.gif\","+l+");this.className=\"DropDownOverStyle\";' onmouseout='clearInterval(objPopCalList["+l+"].intervalID1);PopCalSwapImage(\"popupSuperChangeRight"+_id+"\",\"right1.gif\","+l+");this.className=\"DropDownOutStyle\";window.status=\"\"' onclick='PopCalIncMonth("+l+")' onmousedown='clearTimeout(objPopCalList["+l+"].timeoutID1);objPopCalList["+l+"].timeoutID1=setTimeout(\"PopCalStartIncMonth("+l+")\",100)' onmouseup='clearTimeout(objPopCalList["+l+"].timeoutID1);clearInterval(objPopCalList["+l+"].intervalID1)'>&nbsp;<img id='popupSuperChangeRight"+_id+"' src='"+objPopCal.imgDir+"right1.gif' border='0' />&nbsp;</span>"
			}
			else
			{
				_leftButton="<span id='popupSuperSpanRight"+_id+"' class='DropDownStyle' ondrag='return(false)' onmouseover='PopCalSwapImage(\"popupSuperChangeRight"+_id+"\",\"right2.gif\","+l+");this.className=\"DropDownOverStyle\";' onmouseout='clearInterval(objPopCalList["+l+"].intervalID1);PopCalSwapImage(\"popupSuperChangeRight"+_id+"\",\"right1.gif\","+l+");this.className=\"DropDownOutStyle\";window.status=\"\"' onclick='PopCalDecMonth("+l+")' onmousedown='clearTimeout(objPopCalList["+l+"].timeoutID1);objPopCalList["+l+"].timeoutID1=setTimeout(\"PopCalStartDecMonth("+l+")\",100)' onmouseup='clearTimeout(objPopCalList["+l+"].timeoutID1);clearInterval(objPopCalList["+l+"].intervalID1)'>&nbsp;<img id='popupSuperChangeRight"+_id+"' src='"+objPopCal.imgDir+"right1.gif' border='0' />&nbsp;</span>"
				_rightButton="<span id='popupSuperSpanLeft"+_id+"' class='DropDownStyle' ondrag='return(false)' onmouseover='PopCalSwapImage(\"popupSuperChangeLeft"+_id+"\",\"left2.gif\","+l+");this.className=\"DropDownOverStyle\";' onmouseout='clearInterval(objPopCalList["+l+"].intervalID1);PopCalSwapImage(\"popupSuperChangeLeft"+_id+"\",\"left1.gif\","+l+");this.className=\"DropDownOutStyle\";window.status=\"\"' onclick='PopCalIncMonth("+l+")' onmousedown='clearTimeout(objPopCalList["+l+"].timeoutID1);objPopCalList["+l+"].timeoutID1=setTimeout(\"PopCalStartIncMonth("+l+")\",100)' onmouseup='clearTimeout(objPopCalList["+l+"].timeoutID1);clearInterval(objPopCalList["+l+"].intervalID1)'>&nbsp;<img id='popupSuperChangeLeft"+_id+"' src='"+objPopCal.imgDir+"left1.gif' border='0' />&nbsp;</span>"
			}
			sCalendar="<table id='popupSuperHighLight"+_id+"' style='border:1px solid #a0a0a0;' cellspacing=1 cellpadding=0 ><tr><td style='cursor:default'>"
			sCalendar+="<div class='TitleStyle'><table width='100%'><tr>"
			sCalendar+="<td id='popupSuperCaption"+_id+"'></td>"
			sCalendar+="<td id='popupSuperMoveCalendar"+_id+"' align='center'></td>"
			sCalendar+="<td align='"+((objPopCal.lr==1)?"right":"left")+"' style='cursor:default'>"
			if((objPopCal.buttons==0)||(objPopCal.buttons==2))
			{
				sCalendar+="<span onclick='ImgCloseBoton"+_id+".src=\""+ objPopCal.imgDir+"close.gif\";objPopCalList["+l+"].PopCalTimeOut=window.setTimeout(\"window.clearTimeout(objPopCalList["+l+"].PopCalTimeOut);objPopCalList["+l+"].PopCalTimeOut=null;PopCalHideCalendar("+l+")\",100)'><img id='ImgCloseBoton"+_id+"' src='"+objPopCal.imgDir+"close.gif' onmouseover='if(objPopCalList["+l+"].PopCalDragClose){this.src=\""+ objPopCal.imgDir+"closedown.gif\"}' onmousedown='this.src=\""+ objPopCal.imgDir+"closedown.gif\"' onmouseup='this.src=\""+ objPopCal.imgDir+"close.gif\"' onmouseout='this.src=\""+ objPopCal.imgDir+"close.gif\"' ondrag='objPopCalList["+l+"].PopCalDragClose=true;return(false)' class='CloseButtonStyle' /></span>"
			}
			else
			{
				sCalendar+=_rightButton
			}
			sCalendar+="</td></tr></table></div>"
			sCalendar+="</td></tr>"
			sCalendar+="<tr><td align='center' style='padding:5px;'>"
			sCalendar+="<div id='popupSuperContent"+_id+"' style='white-space:nowrap;'></div>"
			sCalendar+="</td></tr>"
			if(objPopCal.showToday==1)
			{
				sCalendar+="<tr><td style='padding:5px;' class='TodayStyle' align='center'>"
				sCalendar+="<div style='white-space:nowrap;' class='TextStyle' onclick='PopCalChangeCurrentMonth("+l+");'>"+objPopCal.todayString+"</div>"
				sCalendar+="</td></tr>"
			}
			if((objPopCal.BlankFieldText!="")&&(typeof(__PopCalSelectNone)=="function"))
			{
				sCalendar+="<tr><td style='padding:1px;' class='TodayStyle' align='center'>"
				sCalendar+="<div style='white-space:nowrap' class='TextStyle' onclick='__PopCalSelectNone("+l+");'>"+objPopCal.BlankFieldText+"</div>"
				sCalendar+="</td></tr>"
			}
			sCalendar+="</table>"
			_Component.innerHTML=sCalendar
			var sHTML="<nobr>"
			if(objPopCal.buttons!=2)
			{
				sHTML+=_leftButton+((objPopCal.buttons!=3)?"&nbsp;":"")
			}
			if(objPopCal.buttons==0)
			{
				sHTML+=_rightButton
			}
			if(objPopCal.buttons==3)
			{
				sHTML+="<span id='popupSuperSpanMonth"+_id+"' ondrag='return(false)' style='display:none;'></span>"
				sHTML+="<span id='popupSuperSpanYear"+_id+"' ondrag='return(false)' style='display:none;'></span>"
			}
			else
			{
				sHTML+="&nbsp;<span id='popupSuperSpanMonth"+_id+"' class='DropDownStyle' ondrag='return(false)' onmouseover='PopCalSwapImage(\"popupSuperChangeMonth"+_id+"\",\"drop2.gif\","+l+");this.className=\"DropDownOverStyle\";' onmouseout='PopCalSwapImage(\"popupSuperChangeMonth"+_id+"\",\"drop1.gif\","+l+");this.className=\"DropDownOutStyle\";window.status=\"\"' onclick='objPopCalList["+l+"].keepMonth=!PopCalIsObjectVisible(objPopCalList["+l+"].Object.popupSuperMonth);PopCalUpMonth("+l+")'></span>&nbsp;"
				if(objPopCal.buttons!=0)
				{
					sHTML+="&nbsp;"
				}
				sHTML+="<span id='popupSuperSpanYear"+_id+"' class='DropDownStyle' ondrag='return(false)' onmouseover='PopCalSwapImage(\"popupSuperChangeYear"+_id+"\",\"drop2.gif\","+l+");this.className=\"DropDownOverStyle\";' onmouseout='PopCalSwapImage(\"popupSuperChangeYear"+_id+"\",\"drop1.gif\","+l+");this.className=\"DropDownOutStyle\";window.status=\"\"' onclick='objPopCalList["+l+"].keepYear=!PopCalIsObjectVisible(objPopCalList["+l+"].Object.popupSuperYear);PopCalUpYear("+l+")' onmousewheel='PopCalWheelYear("+l+")'></span>&nbsp;"
			}
			sHTML+="</nobr>"
			PopCalGetById("popupSuperCaption"+_id).innerHTML=sHTML
			PopCalGetById("popupSuperCaption"+_id).dropDown=(objPopCal.buttons!=3)
			if(objPopCal.ie)
			{
				if(objPopCal.move==1)
				{
					var superMoveCalendar=PopCalGetById("popupSuperMoveCalendar"+_id)
					superMoveCalendar.width="100%"
					superMoveCalendar.onmousedown=new Function("PopCalDrag("+l+")")
					superMoveCalendar.ondblclick=new Function("PopCalMoveDefault("+l+")")
					superMoveCalendar.onmouseup=new Function("PopCalDrop("+l+")")
				}
			}
			else
			{
				objPopCal.keepInside=0
			}
			PopCal.startAt=objPopCal.startAt
			PopCal.weekend=objPopCal.weekend
			PopCal.clientValidator=objPopCal.clientValidator
			PopCal.showWeekNumber=objPopCal.showWeekNumber
			PopCal.weekNumberFormula=objPopCal.weekNumberFormula
			PopCal.showToday=objPopCal.showToday
			PopCal.showDaysOutMonth=objPopCal.showDaysOutMonth
			PopCal.showWeekend=objPopCal.showWeekend
			PopCal.showHolidays=objPopCal.showHolidays
			PopCal.showSpecialDay=objPopCal.showSpecialDay
			PopCal.selectWeekend=objPopCal.selectWeekend
			PopCal.selectHoliday=objPopCal.selectHoliday
			PopCal.addCarnival=objPopCal.addCarnival
			PopCal.addGoodFriday=objPopCal.addGoodFriday
			PopCal.defaultFormat=objPopCal.defaultFormat
			PopCal.fixedX=objPopCal.fixedX
			PopCal.fixedY=objPopCal.fixedY
			PopCal.fade=objPopCal.fade
			PopCal.shadow=objPopCal.shadow
			PopCal.centuryLimit=objPopCal.centuryLimit
			PopCal.move=objPopCal.move
			PopCal.saveMovePos=objPopCal.saveMovePos
			PopCal.keepInside=objPopCal.keepInside
			PopCal.saveKeepInside=objPopCal.keepInside
			PopCal.popupSuperCalendar=PopCalGetById("popupSuperCalendar"+_id)
			PopCal.popupSuperShadowRight=PopCalGetById("popupSuperShadowRight")
			PopCal.popupSuperShadowBottom=PopCalGetById("popupSuperShadowBottom")
			PopCal.popupSuperMonth=PopCalGetById("popupSuperMonth")
			PopCal.popupSuperYear=PopCalGetById("popupSuperYear")
			PopCal.popupSuperYearList=[]
			PopCal.popupSuperCalendar.OverSelect=PopCalGetById("popupOverCalendar")
			PopCal.popupSuperMonth.OverSelect=PopCalGetById("popupOverYearMonth")
			PopCal.popupSuperYear.OverSelect=PopCalGetById("popupOverYearMonth")
			if(objPopCal.ie)
			{
				if(PopCal.shadow==1)
				{
					PopCal.popupSuperCalendar.ShadowOverSelect=PopCalGetById("popupOverShadow")
					PopCal.popupSuperCalendar.lr=objPopCal.lr
				}
				PopCal.popupSuperCalendar.style.filter="blendTrans()"
				PopCal.popupSuperShadowRight.style.filter="alpha(opacity=50)"
				PopCal.popupSuperShadowBottom.style.filter="alpha(opacity=50)"
				if((objPopCal.ieVersion<5.5)||(typeof(PopCalGetById("CalendarLoadFilters").filters)!="object"))
				{
					PopCal.fade=0
				}
			}
			else
			{
				PopCal.popupSuperShadowRight.style.MozOpacity=.5
				PopCal.popupSuperShadowBottom.style.MozOpacity=.5
				if(typeof(PopCal.popupSuperCalendar.style.MozOpacity)!='string')
				{
					PopCal.fade=0
				}
			}
			if(PopCal.fade<0) PopCal.fade=0
			if(PopCal.fade>1) PopCal.fade=1
			if(objPopCal.lr==1)
			{
				PopCal.popupSuperCalendar.dir="ltr"
				PopCal.popupSuperMonth.dir="ltr"
				PopCal.popupSuperYear.dir="ltr"
			}
			else
			{
				PopCal.popupSuperCalendar.dir="rtl"
				PopCal.popupSuperMonth.dir="rtl"
				PopCal.popupSuperYear.dir="rtl"
			}
			PopCal.initialized=1
		}
	}
}

function PopCalRightToLeft()
{
	var _obj=document.getElementsByTagName("BODY")
	if(_obj.length>=1)
	{
		if(_obj[0].dir.toLowerCase()=="rtl") return(true)
		if(_obj[0].dir!="") return(false)
	}
	_obj=document.getElementsByTagName("HTML")
	if(_obj.length>=1)
	{
		if(_obj[0].dir.toLowerCase()=="rtl") return(true)
	}
	return(false)
}

function PopCalCalendarVisible()
{
	for(var i=0;i<objPopCalList.length;i++)
	{
		if(objPopCalList[i].Object.popupSuperCalendar.style.visibility!="hidden")
		{
			return(objPopCalList[i])
		}
	}
	return(null)
}

function PopCalSetFocus(ctl)
{
	try
	{
		ctl.focus()
	}
	catch(e)
	{
		//Nothing
	}
}

function PopCalSetPosition(o,t,l,h,w)
{
	if(t) o.style.top=t+'px'
	if(l) o.style.left=l+'px'
	if(h) o.style.height=h+'px'
	if(w) o.style.width=w+'px'
	if(o.ShadowOverSelect)
	{
		o.ShadowOverSelect.style.top=(parseInt(o.style.top,10)+10)+'px'
		if(o.lr==1)
		{
			o.ShadowOverSelect.style.left=(parseInt(o.style.left,10)+10)+'px'
		}
		else
		{
			o.ShadowOverSelect.style.left=(parseInt(o.style.left,10)-10)+'px'
		}
		o.ShadowOverSelect.style.height=(o.offsetHeight+3)+'px'
		o.ShadowOverSelect.style.width=(o.offsetWidth)+'px'
	}
	if(o.OverSelect)
	{
		o.OverSelect.style.top=(parseInt(o.style.top,10))+'px'
		o.OverSelect.style.left=(parseInt(o.style.left,10))+'px'
		o.OverSelect.style.height=(o.offsetHeight)+'px'
		o.OverSelect.style.width=(o.offsetWidth)+'px'
	}
}

function PopCalShow(ctl,format,from,to,execute,l)
{
	var objPopCal=objPopCalList[l]
	var PopCal=objPopCal.Object
	if(PopCal)
	{
		if(PopCal.initialized==1)
		{
			if(document.body)
			{
				PopCalSetFocus(document.body)
			}
			objPopCal.movePopCal=false
			if(objPopCal.timeoutID3)
			{
				clearTimeout(objPopCal.timeoutID3)
				objPopCal.timeoutID3=null
			}
			var objPopCalVisible=PopCalCalendarVisible()
			if(objPopCalVisible==null)
			{
				objPopCal.commandExecute=null
				if(execute)
				{
					objPopCal.commandExecute=execute
				}
				if(objPopCal.ie)
				{
					objPopCal.onKeyPress=document.onkeyup
					document.onkeyup=new Function("objPopCalList["+l+"].bShow=false;PopCalSetFocus(objPopCalList["+l+"].ctl());PopCalClickDocumentBody("+l+");")
					objPopCal.onmouseup=document.onmouseup
					document.onmouseup=new Function("objPopCalList["+l+"].movePopCal=false;if(event.button==2){objPopCalList["+l+"].bShow=false;PopCalClickDocumentBody("+l+");}")
					if(PopCal.move==1)
					{
						objPopCal.onmousemove=document.onmousemove
						document.onmousemove=new Function('PopCalTrackMouse('+l+');')
					}
					objPopCal.onresize=window.onresize
					window.onresize=new Function('PopCalScroll('+l+');')
					if(PopCal.keepInside==1)
					{
						objPopCal.onscroll=window.onscroll
						window.onscroll=new Function('PopCalScroll('+l+');')
					}
				}
				else
				{
					objPopCal.onKeyPress=document.onkeyup
					document.captureEvents(Event.KEYUP)
					document.onkeyup=new Function("objPopCalList["+l+"].bShow=false;PopCalClickDocumentBody("+l+");PopCalSetFocus(objPopCalList["+l+"].ctl());")
				}
				objPopCal.onClick=document.onclick
				document.onclick=new Function('PopCalClickDocumentBody('+l+');')
				if(objPopCal.ie)
				{
					objPopCal.onSelectStart=document.onselectstart
					document.onselectstart=new Function('return(false);')
					objPopCal.onContextMenu=document.oncontextmenu
					document.oncontextmenu=new Function('return(false);')
				}
				objPopCal.yearConstructed=false
				objPopCal.monthConstructed=false
				objPopCal.ctlId=ctl.id
				PopCalSetScroll("Div",l,false)
				objPopCal.dateFormat=""
				if(format)
				{
					objPopCal.dateFormat=format.toLowerCase()
				}
				else if(PopCal.defaultFormat)
				{
					objPopCal.dateFormat=PopCal.defaultFormat.toLowerCase()
				}
				objPopCal.dateFrom=01
				objPopCal.monthFrom=00
				objPopCal.yearFrom=1900
				objPopCal.dateUpTo=31
				objPopCal.monthUpTo=11
				objPopCal.yearUpTo=2099
				objPopCal.dateSelected=0
				objPopCal.monthSelected=objPopCal.monthNow
				objPopCal.yearSelected=objPopCal.yearNow
				if(PopCalSetDMY(ctl.value,objPopCal.dateFormat,l))
				{
					objPopCal.dateSelected=objPopCal.oDate
					objPopCal.monthSelected=objPopCal.oMonth
					objPopCal.yearSelected=objPopCal.oYear
				}
				if(from)
				{
					if(PopCalIsToday(from))
					{
						objPopCal.dateFrom=objPopCal.today.getDate()
						objPopCal.monthFrom=objPopCal.today.getMonth()
						objPopCal.yearFrom=objPopCal.today.getFullYear()
					}
					else if(PopCalSetDMY(from,objPopCal.dateFormat,l))
					{
						objPopCal.dateFrom=objPopCal.oDate
						objPopCal.monthFrom=objPopCal.oMonth
						objPopCal.yearFrom=objPopCal.oYear
					}
				}
				if(to)
				{
					if(PopCalIsToday(to))
					{
						objPopCal.dateUpTo=objPopCal.today.getDate()
						objPopCal.monthUpTo=objPopCal.today.getMonth()
						objPopCal.yearUpTo=objPopCal.today.getFullYear()
					}
					else if(PopCalSetDMY(to,objPopCal.dateFormat,l))
					{
						objPopCal.dateUpTo=objPopCal.oDate
						objPopCal.monthUpTo=objPopCal.oMonth
						objPopCal.yearUpTo=objPopCal.oYear
					}
				}
				if(!PopCalCenturyOn(objPopCal.dateFormat))
				{
					if(PopCalDateFrom(l)<PopCalPad(1900+objPopCal.centuryLimit,4,"0","L")+"0001")
					{
						objPopCal.dateFrom=01
						objPopCal.monthFrom=00
						objPopCal.yearFrom=1900+objPopCal.centuryLimit
					}
					if(PopCalDateUpTo(l)>PopCalPad(2000+(objPopCal.centuryLimit-1),4,"0","L")+"1131")
					{
						objPopCal.dateUpTo=31
						objPopCal.monthUpTo=11
						objPopCal.yearUpTo=2000+(objPopCal.centuryLimit-1)
					}
				}
				if(PopCalDateFrom(l)>PopCalDateUpTo(l))
				{
					objPopCal.oDate=objPopCal.dateFrom
					objPopCal.oMonth=objPopCal.monthFrom
					objPopCal.oYear=objPopCal.yearFrom
					objPopCal.dateFrom=objPopCal.dateUpTo
					objPopCal.monthFrom=objPopCal.monthUpTo
					objPopCal.yearFrom=objPopCal.yearUpTo
					objPopCal.dateUpTo=objPopCal.oDate
					objPopCal.monthUpTo=objPopCal.oMonth
					objPopCal.yearUpTo=objPopCal.oYear
				}
				if(PopCalDateSelect(l)<PopCalDateFrom(l))
				{
					objPopCal.dateSelected=0
					objPopCal.monthSelected=objPopCal.monthFrom
					objPopCal.yearSelected=objPopCal.yearFrom
				}
				if(PopCalDateSelect(l)>PopCalDateUpTo(l))
				{
					objPopCal.dateSelected=0
					objPopCal.monthSelected=objPopCal.monthUpTo
					objPopCal.yearSelected=objPopCal.yearUpTo
				}
				objPopCal.odateSelected=objPopCal.dateSelected
				objPopCal.omonthSelected=objPopCal.monthSelected
				objPopCal.oyearSelected=objPopCal.yearSelected
				PopCalMoveDefaultPos(l)
				if(objPopCal.ie)
				{
					if((PopCal.move==1)&&(PopCal.saveMovePos==1))
					{
						if(objPopCal.ctl())
						{
							if(objPopCal.ctl().CalendarTop)
							{
								PopCalSetPosition(PopCal.popupSuperCalendar,objPopCal.ctl().CalendarTop)
							}
							if(objPopCal.ctl().CalendarLeft)
							{
								PopCalSetPosition(PopCal.popupSuperCalendar,null,objPopCal.ctl().CalendarLeft)
							}
						}
					}
				}
				PopCalConstructCalendar(l)
				PopCalFadeIn(l)
				PopCalScroll(l)
				objPopCal.bShow=true
			}
			else
			{
				objPopCalVisible.executeFade=(objPopCalVisible.ctlIdNow==ctl.id)
				objPopCal.executeFade=(objPopCalVisible.ctlIdNow==ctl.id)
				PopCalHideCalendar(objPopCalVisible.id)
				if(objPopCalVisible.ctl())
				{
					objPopCalVisible.ctlId="n:1"
				}
				if(objPopCal!=objPopCalVisible)
				{
					objPopCal.ctlIdNow="n:2"
				}
				if(objPopCal.ctlIdNow!=ctl.id)
				{
					PopCalShow(ctl,format,from,to,execute,objPopCal.id)
				}
				objPopCal.executeFade=true
				objPopCalVisible.executeFade=true
			}
			objPopCal.ctlIdNow=ctl.id
		}
	}
}

function PopCalAddDays(dateValue,format,daysToAdd,l)
{
	var objPopCal=objPopCalList[l]
	if((dateValue)&&(dateValue!=""))
	{
		var sDateFormat=((format==null)?objPopCal.Object.defaultFormat.toLowerCase():format.toLowerCase())
		var incDays=((daysToAdd==null)?0:daysToAdd)
		var dFecha=null
		if(PopCalIsToday(dateValue))
		{
			dFecha=PopCalSetDays(objPopCal.today, incDays)
		}
		else if(PopCalSetDMY(dateValue,sDateFormat,l))
		{
			dFecha=PopCalSetDays(PopCalGetDate(dateValue,sDateFormat,l), incDays)
		}
		if(dFecha) return(PopCalConstructDate(dFecha.getDate(),dFecha.getMonth(),dFecha.getFullYear(),sDateFormat,l))
	}
	return("")
}

function PopCalScroll(l)
{
	var objPopCal=objPopCalList[l]
	var objCal=objPopCal.Object.popupSuperCalendar
	var obj=objCal.OverSelect
	if(obj)
	{
		obj.style.visibility='hidden'
		obj.style.visibility='visible'
	}
	obj=objCal.ShadowOverSelect
	if(obj)
	{
		obj.style.visibility='hidden'
		obj.style.visibility='visible'
	}
	if(objCal.style.visibility!="hidden")
	{
		if((objPopCal.ctl().CalendarTop==null)&&(objPopCal.ctl().CalendarLeft==null))
		{
			PopCalDownMonth(l)
			PopCalDownYear(l)
			PopCalMoveDefault(l)
		}
	}
}

function PopCalMoveDefaultPos(l)
{
	var objPopCal=objPopCalList[l]
	var PopCal=objPopCal.Object
	var leftpos=0
	var toppos=0
	var lDivTop=-1
	var objCal=PopCal.popupSuperCalendar
	var mTop=0
	var mLeft=0
	var mRight=0
	var mBottom=0
	var KeepInside=true
	if(((PopCal.fixedX==-1)&&(PopCal.fixedY==-1)&&(objPopCal.ctl().style.display!='none'))||(objPopCal.ControlAlignLeft))
	{
		var aTag=null
		if((objPopCal.lr==0)&&(objPopCal.ControlAlignLeft==null))
		{
			objPopCal.ControlAlignLeft=objPopCal.ctl()
		}
		if(objPopCal.ControlAlignLeft)
		{
			aTag=objPopCal.ControlAlignLeft
		}
		else
		{
			aTag=objPopCal.ctl()
		}
		if(aTag.style.position.toLowerCase()!='absolute')
		{
			if(document.body)
			{
				if(document.body.offsetTop!=0)
				{
					KeepInside=false
					if(document.body.currentStyle)
					{
						if(document.body.currentStyle.marginTop) mTop=parseInt(document.body.currentStyle.marginTop,10)
						if(document.body.currentStyle.marginLeft) mLeft=parseInt(document.body.currentStyle.marginLeft,10)
						if(document.body.currentStyle.marginRight) mRight=parseInt(document.body.currentStyle.marginRight,10)
						if(document.body.currentStyle.marginBottom) mBottom=document.body.currentStyle.marginBottom
					}
				}
			}
		}
		leftpos+=aTag.offsetLeft
		toppos+=aTag.offsetTop
		aTag=aTag.offsetParent
		while((aTag.tagName!="BODY")&&(aTag.tagName!="HTML"))
		{
			leftpos+=aTag.offsetLeft
			toppos+=aTag.offsetTop
			if(aTag.tagName=="DIV")
			{
				if(lDivTop==-1)
				{
					lDivTop+=(1+aTag.offsetTop)
				}
				leftpos-=aTag.scrollLeft
				toppos-=aTag.scrollTop
			}
			else if(lDivTop!=-1)
			{
				lDivTop+=aTag.offsetTop
			}
			aTag=aTag.offsetParent
		}
	}
	else
	{
		var aTag=document.body
	}
	if(objPopCal.ControlAlignLeft)
	{
		leftpos+=objPopCal.ControlAlignLeft.offsetWidth-objCal.offsetWidth
		toppos+=objPopCal.ControlAlignLeft.offsetHeight+7
	}
	else
	{
		leftpos=PopCal.fixedX==-1?leftpos:PopCal.fixedX
		toppos=PopCal.fixedY==-1?toppos+objPopCal.ctl().offsetHeight+7:PopCal.fixedY
	}
	leftpos+=mLeft
	toppos+=mTop
	if(objPopCal.ie)
	{
		if(PopCalRightToLeft())
		{
			PopCal.keepInside=0
		}
		else
		{
			PopCal.keepInside=PopCal.saveKeepInside
		}
	}
	if((PopCal.keepInside==1)&&(KeepInside))
	{
		if(((leftpos+objCal.offsetWidth+10+((PopCal.shadow==1)?25:15))-aTag.scrollLeft)>(mLeft+aTag.offsetWidth+mRight))
		{
			leftpos-=(((((leftpos+objCal.offsetWidth)-(mLeft+aTag.offsetWidth+mRight))+10)-aTag.scrollLeft)+((PopCal.shadow==1)?25:15))
		}
		if(leftpos<aTag.scrollLeft+10)
		{
			leftpos=aTag.scrollLeft+10
		}
		if(toppos<lDivTop)
		{
			toppos=lDivTop
		}
		if(((toppos+objCal.offsetHeight+85)-aTag.scrollTop)>(mTop+aTag.offsetHeight+mBottom))
		{
			toppos-=((((toppos+objCal.offsetHeight)-(mTop+aTag.offsetHeight+mBottom))+75)-aTag.scrollTop)
		}
		if(toppos<aTag.scrollTop+10)
		{
			toppos=aTag.scrollTop+10
		}
		if(leftpos<10)
		{
			leftpos=10
		}
	}
	PopCalSetPosition(objCal,toppos,leftpos)
}

function PopCalMoveDefault(l)
{
	var objPopCal=objPopCalList[l]
	var PopCal=objPopCal.Object
	PopCalMoveDefaultPos(l)
	PopCalMoveShadow(l)
	if(PopCal.saveMovePos==1)
	{
		if(objPopCal.ctl())
		{
			objPopCal.ctl().CalendarLeft=null
			objPopCal.ctl().CalendarTop=null
		}
	}
	objPopCal.bShow=false
}

function PopCalDrag(l)
{
	var objPopCal=objPopCalList[l]
	if((!objPopCal.movePopCal)&&(event.button==1))
	{
		var PopCal=objPopCal.Object
		var ex=event.clientX+document.body.scrollLeft
		var ey=event.clientY+document.body.scrollTop
		PopCalGetById("popupSuperHighLight"+PopCal.id).style.borderColor="red"
		var obj=PopCal.popupSuperCalendar
		obj.style.xoff=parseInt(obj.style.left,10)-ex
		obj.style.yoff=parseInt(obj.style.top,10)-ey
		PopCalDownMonth(l)
		PopCalDownYear(l)
		objPopCal.movePopCal=true
	}
	objPopCal.bShow=true
}

function PopCalTrackMouse(l)
{
	var objPopCal=objPopCalList[l]
	var PopCal=objPopCal.Object
	var obj=PopCal.popupSuperCalendar
	if(objPopCal.movePopCal)
	{
		if(event.button!=1)
		{
			PopCalDrop(l)
			objPopCal.bShow=true
			return
		}
		var x=event.clientX+document.body.scrollLeft
		var y=event.clientY+document.body.scrollTop
		var lLeft=(obj.style.xoff+x)
		var lTop=(obj.style.yoff+y)
		if((parseInt(obj.style.left,10)!=parseInt(lLeft,10))||(parseInt(obj.style.top,10)!=parseInt(lTop,10)))
		{
			PopCalSetPosition(obj,lTop,lLeft)
			x=event.clientX+document.body.scrollLeft
			y=event.clientY+document.body.scrollTop
			PopCalMoveShadow(l)
			if(PopCal.saveMovePos==1)
			{
				if(objPopCal.ctl())
				{
					objPopCal.ctl().CalendarLeft=parseInt(obj.style.left,10)
					objPopCal.ctl().CalendarTop=parseInt(obj.style.top,10)
				}
			}
		}
		objPopCal.bShow=true
	}
}

function PopCalDrop(l)
{
	var objPopCal=objPopCalList[l]
	objPopCal.bShow=true
	objPopCal.movePopCal=false
	PopCalGetById("popupSuperHighLight"+objPopCal.id).style.borderColor="#a0a0a0"
}

function PopCalValidateType1(_date,_Type1)
{
	if(parseInt(_Type1.d,10)==_date.getDate())
	{
		if((parseInt(_Type1.m,10)==0)||((parseInt(_Type1.m,10)==(_date.getMonth()+1))&&(parseInt(_Type1.m,10)!=0)))
		{
			if((parseInt(_Type1.y,10)==0)||((parseInt(_Type1.y,10)==_date.getFullYear())&&(parseInt(_Type1.y,10)!=0)))
			{
				return(true)
			}
		}
	}
	return(false)
}

function PopCalHolidayRec1(d,m,y,t,tipo)
{
	this.d=d
	this.m=m
	this.y=y
	this.desc=t
	this.tipo=tipo
	this.type="Type 1"
}

function PopCalValidateType2(_date,_Type2)
{
	var _NewDate
	if((_date.getDay()==_Type2.dw)&&((_date.getMonth()==(_Type2.m-1))||(_Type2.m==0)))
	{
		if(_Type2.s==0)
		{
			return(true)
		}
		else if(_Type2.s==-1)
		{
			_NewDate=PopCalSetDays(_date,7)
			if(_date.getMonth()!=_NewDate.getMonth())
			{
				return(true)
			}
		}
		else
		{
			for(var i=1;i<=5;i++)
			{
				_NewDate=PopCalSetDays(_date,-7*i)
				if(_date.getMonth()!=_NewDate.getMonth())
				{
					break
				}
			}
			if(i==_Type2.s)
			{
				return(true)
			}
		}
	}
	return(false)
}

function PopCalHolidayRec2(s,dw,m,t,tipo)
{
	this.s=s
	this.dw=dw
	this.m=m
	this.desc=t
	this.tipo=tipo
	this.type="Type 2"
}

function PopCalValidateType3(_date,_Type3)
{
	var BeginDate=new Date(_Type3.y,_Type3.m-1,_Type3.d)
	if(BeginDate<=_date)
	{
		var Interval=_Type3.i*(((_Type3.f==1)?7:1)*86400000)
		if((((_date-BeginDate)%Interval)==0))
		{
			if(_Type3.r==0)
			{
				return(true)
			}
			else
			{
				return(((_date-BeginDate) / Interval)<_Type3.r)
			}
		}
	}
	return(false)
}

function PopCalHolidayRec3(d,m,y,i,f,r,t,tipo)
{
	this.d=d
	this.m=m
	this.y=y
	this.i=i
	this.f=f
	this.r=r
	this.desc=t
	this.tipo=tipo
	this.type="Type 3"
}

function PopCalAddHoliday(d,m,y,t,l)
{
	var objPopCal=objPopCalList[l]
	objPopCal.Holidays[objPopCal.HolidaysCounter++]=new PopCalHolidayRec1(d,m,y,t,1)
}

function PopCalAddSpecialDay(d,m,y,t,l)
{
	var objPopCal=objPopCalList[l]
	objPopCal.Holidays[objPopCal.HolidaysCounter++]=new PopCalHolidayRec1(d,m,y,t,0)
}

function PopCalAddIrregularHoliday(s,dw,m,t,l)
{
	var objPopCal=objPopCalList[l]
	objPopCal.Holidays[objPopCal.HolidaysCounter++]=new PopCalHolidayRec2(s,dw,m,t,1)
}

function PopCalAddIrregularSpecialDay(s,dw,m,t,l)
{
	var objPopCal=objPopCalList[l]
	objPopCal.Holidays[objPopCal.HolidaysCounter++]=new PopCalHolidayRec2(s,dw,m,t,0)
}

function PopCalAddRecurrenceSpecialDay(d,m,y,i,f,r,t,l)
{
	var objPopCal=objPopCalList[l]
	objPopCal.Holidays[objPopCal.HolidaysCounter++]=new PopCalHolidayRec3(d,m,y,i,f,r,t,0)
}

function PopCalFormatDate(dateValue,oldFormat,newFormat,l)
{
	var objPopCal=objPopCalList[l]
	var PopCal=objPopCal.Object
	var newValue=""
	if(PopCal)
	{
		var formatOld=PopCal.defaultFormat
		if(oldFormat)
		{
			formatOld=oldFormat
		}
		var formatNew=PopCal.defaultFormat
		if(newFormat)
		{
			formatNew=newFormat
		}
		if(PopCalIsToday(dateValue))
		{
			newValue=PopCalConstructDate(objPopCal.today.getDate(),objPopCal.today.getMonth(),objPopCal.today.getFullYear(),formatNew,l)
		}
		else if(PopCalSetDMY(dateValue,formatOld,l))
		{
			var _Date=new Date(objPopCal.oYear,objPopCal.oMonth,objPopCal.oDate)
			if((_Date.getDate()==objPopCal.oDate)&&(_Date.getMonth()==objPopCal.oMonth)&&(_Date.getFullYear()==objPopCal.oYear))
			{
				newValue=PopCalConstructDate(objPopCal.oDate,objPopCal.oMonth,objPopCal.oYear,formatNew,l)
			}
		}
	}
	return(newValue)
}

function PopCalForcedToday(dateValue,format,l)
{
	var objPopCal=objPopCalList[l]
	if(objPopCal.Object)
	{
		objPopCal.forceTodayTo=dateValue
		objPopCal.forceTodayFormat=format
	}
}

function PopCalSetScroll(elmID,l,restore)
{
	var objPopCal=objPopCalList[l]
	if(objPopCal.ie)
	{
		if(objPopCal.ctl())
		{
			if(PopCalIsObjectVisible(objPopCal.ctl()))
			{
				var objParent=objPopCal.ctl().offsetParent
				while((objParent)&&(objParent.tagName!="BODY")&&(objParent.tagName!="HTML"))
				{
					if(objParent.tagName.toLowerCase()==elmID.toLowerCase())
					{
						if(restore)
						{
							objParent.onscroll=null
							if(objParent.savedScroll) objParent.onscroll=objParent.savedScroll
							objParent.savedScroll=null
						}
						else
						{
							objParent.savedScroll=objParent.onscroll
							objParent.onscroll=new Function("PopCalScroll("+l+");")
						}
					}
					objParent=objParent.offsetParent
				}
			}
		}
	}
}

function PopCalSwapImage(srcImg,destImg,l)
{
	var objPopCal=objPopCalList[l]
	PopCalGetById(srcImg).src=objPopCal.imgDir+destImg
}

function PopCalHideCalendar(l,HideNow)
{
	var objPopCal=objPopCalList[l]
	if(!objPopCal)
	{
		objPopCal=null
		return(false)
	}
	var PopCal=objPopCal.Object
	if(PopCal.popupSuperCalendar.style.visibility!="hidden")
	{
		PopCalSetScroll("Div",l,true)
		if(objPopCal.ie)
		{
			document.onkeyup=objPopCal.onKeyPress
		}
		else
		{
			document.releaseEvents(Event.KEYUP)
			document.onkeyup=objPopCal.onKeyPress
		}
		document.onclick=objPopCal.onClick
		if(objPopCal.ie)
		{
			document.onselectstart=objPopCal.onSelectStart
			document.oncontextmenu=objPopCal.onContextMenu
		}
		if(objPopCal.ie)
		{
			if(PopCal.move==1)
			{
				document.onmousemove=objPopCal.onmousemove
			}
			document.onmouseup=objPopCal.onmouseup
			window.onresize=objPopCal.onresize
			if(PopCal.keepInside==1)
			{
				window.onscroll=objPopCal.onscroll
			}
		}
		objPopCal.onKeyPress=null
		objPopCal.onClick=null
		objPopCal.onSelectStart=null
		objPopCal.onContextMenu=null
		objPopCal.onmousemove=null
		objPopCal.onmouseup=null
		objPopCal.onresize=null
		objPopCal.onscroll=null
		if(PopCal.popupSuperMonth)
		{
			PopCal.popupSuperMonth.style.display="none"
			if(PopCal.popupSuperMonth.OverSelect)
			{
				PopCal.popupSuperMonth.OverSelect.style.display="none"
				PopCal.popupSuperMonth.OverSelect.style.height='1px'
				PopCal.popupSuperMonth.OverSelect.style.width='1px'
			}
		}
		if(PopCal.popupSuperYear)
		{
			PopCal.popupSuperYear.style.display="none"
			if(PopCal.popupSuperYear.OverSelect)
			{
				PopCal.popupSuperYear.OverSelect.style.display="none"
				PopCal.popupSuperYear.OverSelect.style.height='1px'
				PopCal.popupSuperYear.OverSelect.style.width='1px'
			}
		}
		PopCalFadeOut(l,HideNow)
		if(!objPopCal)
		{
			objPopCal=null
			return(false)
		}
	}
}

function PopCalMozFadeIn(l)
{
	var objCal=objPopCalList[l].Object.popupSuperCalendar
	if((parseFloat(objCal.style.MozOpacity)+objCal.Opacity)>=.99)
	{
		objCal.style.MozOpacity=.99
		clearInterval(objCal.MozFadeInInterval)
		objCal.MozFadeInInterval=null
	}
	else
	{
		objCal.style.MozOpacity=(parseFloat(objCal.style.MozOpacity)+objCal.Opacity)
	}
}

function PopCalFadeIn(l)
{
	var objPopCal=objPopCalList[l]
	var PopCal=objPopCal.Object
	var objCal=PopCal.popupSuperCalendar
	var objShdR=PopCal.popupSuperShadowRight
	var objShdB=PopCal.popupSuperShadowBottom
	var objOver=objCal.OverSelect
	var objOverShadow=objCal.ShadowOverSelect
	if(objCal.MozFadeOutInterval)
	{
		clearInterval(objCal.MozFadeOutInterval)
		objCal.MozFadeOutInterval=null
	}
	if((PopCal.fade>0)&&(objPopCal.executeFade))
	{
		if(objPopCal.ie)
		{
			objCal.filters.blendTrans.Stop()
			objCal.style.filter="blendTrans(duration="+PopCal.fade+")"
			if((objCal.style.visibility!="visible")&&(objCal.filters.blendTrans.status!=2))
			{
				if(PopCal.shadow==1)
				{
					objShdR.style.filter="alpha(opacity=50)"
					objShdB.style.filter="alpha(opacity=50)"
				}
				objCal.filters.blendTrans.Apply()
				objCal.style.visibility="visible"
				objCal.filters.blendTrans.Play()
				if(PopCal.shadow==1)
				{
					objShdR.style.visibility="visible"
					objShdB.style.visibility="visible"
				}
			}
			else
			{
				if(PopCal.shadow==1)
				{
					objShdR.style.visibility="visible"
					objShdB.style.visibility="visible"
				}
				objCal.style.visibility="visible"
			}
		}
		else
		{
			if(PopCal.shadow==1)
			{
				objShdR.style.display="none"
				objShdR.style.visibility="visible"
				objShdR.style.MozOpacity=.5
				objShdR.style.display=""
				objShdB.style.display="none"
				objShdB.style.visibility="visible"
				objShdB.style.MozOpacity=.5
				objShdB.style.display=""
			}
			objCal.style.display="none"
			objCal.style.visibility="visible"
			objCal.style.display=""
			if(!objCal.MozFadeInInterval)
			{
				objCal.Opacity=(1/(PopCal.fade*10))
				objCal.style.MozOpacity=0
				objCal.MozFadeInInterval=setInterval("PopCalMozFadeIn("+l+")", 50)
			}
			else
			{
				clearInterval(objCal.MozFadeInInterval)
				objCal.MozFadeInInterval=null
				objCal.style.MozOpacity=.99
			}
		}
	}
	else
	{
		if(PopCal.shadow==1)
		{
			objShdR.style.visibility="visible"
			objShdR.style.filter="alpha(opacity=50)"
			objShdR.style.MozOpacity=.5
			objShdB.style.visibility="visible"
			objShdB.style.filter="alpha(opacity=50)"
			objShdB.style.MozOpacity=.5
		}
		objCal.style.visibility="visible"
		objCal.style.MozOpacity=.99
		if(objCal.MozFadeInInterval)
		{
			clearInterval(objCal.MozFadeInInterval)
			objCal.MozFadeInInterval=null
		}
	}
	if(objOver) objOver.style.display=''
	if(objOverShadow) objOverShadow.style.display=''
}

function PopCalMozFadeOut(l)
{
	var objCal=objPopCalList[l].Object.popupSuperCalendar
	if((parseFloat(objCal.style.MozOpacity)-objCal.Opacity)<=0)
	{
		objCal.style.MozOpacity=0
		objCal.style.visibility="hidden"
		clearInterval(objCal.MozFadeOutInterval)
		objCal.MozFadeOutInterval=null
		PopCalMoveTo(0,0,l)
	}
	else
	{
		objCal.style.MozOpacity=(parseFloat(objCal.style.MozOpacity)-objCal.Opacity)
	}
}

function PopCalFadeOut(l,HideNow)
{
	var objPopCal=objPopCalList[l]
	var PopCal=objPopCal.Object
	var objCal=PopCal.popupSuperCalendar
	var objShdR=PopCal.popupSuperShadowRight
	var objShdB=PopCal.popupSuperShadowBottom
	if(objCal.MozFadeInInterval)
	{
		clearInterval(objCal.MozFadeInInterval)
		objCal.MozFadeInInterval=null
	}
	if((objPopCal.ie)&&(PopCal.fade>0)&&(objPopCal.executeFade)&&(!HideNow))
	{
		objCal.filters.blendTrans.Stop()
		objCal.style.filter="blendTrans(duration="+PopCal.fade+")"
		if((objCal.style.visibility!="hidden")&&(objCal.filters.blendTrans.status!=2))
		{
			if(PopCal.shadow==1)
			{
				objShdR.style.filter="alpha(opacity=2)"
				objShdB.style.filter="alpha(opacity=2)"
			}
			objCal.filters.blendTrans.Apply()
			objCal.style.visibility="hidden"
			objCal.filters.blendTrans.Play()
			objPopCal.timeoutID3=setTimeout("PopCalMoveTo(0,0,"+l+")",(PopCal.fade+.05)*1000)
		}
		else
		{
			objCal.style.visibility="hidden"
			PopCalMoveTo(0,0,l)
		}
	}
	else if((!objPopCal.ie)&&(PopCal.fade>0)&&(objPopCal.executeFade)&&(!HideNow)&&(!objCal.MozFadeOutInterval))
	{
		if(typeof(objCal.style.MozOpacity)=='string')
		{
			if(PopCal.shadow==1)
			{
				objShdR.style.MozOpacity=.02
				objShdB.style.MozOpacity=.02
			}
			objCal.Opacity=(1/(PopCal.fade*10))
			objCal.style.MozOpacity=.99
			objCal.MozFadeOutInterval=setInterval("PopCalMozFadeOut("+l+")", 75)
		}
		else
		{
			objCal.style.visibility="hidden"
			PopCalMoveTo(0,0,l)
		}
	}
	else
	{
		objCal.style.visibility="hidden"
		PopCalMoveTo(0,0,l)
	}
}

function PopCalMoveTo(x,y,l)
{
	if(!objPopCalList) return(true)
	var objPopCal=objPopCalList[l]
	if(!objPopCal) return(true)
	if(PopCalCalendarVisible()==null)
	{
		var PopCal=objPopCal.Object
		var objCal=PopCal.popupSuperCalendar
		var objShdR=PopCal.popupSuperShadowRight
		var objShdB=PopCal.popupSuperShadowBottom
		var objOver=objCal.OverSelect
		var objOverShadow=objCal.ShadowOverSelect
		if(objCal.MozFadeOutInterval)
		{
			clearInterval(objCal.MozFadeOutInterval)
			objCal.MozFadeOutInterval=null
		}
		objCal.style.left=x+'px'
		objCal.style.top=y+'px'
		if(PopCal.shadow==1)
		{
			objShdR.style.filter="alpha(opacity=50)"
			objShdB.style.filter="alpha(opacity=50)"
			objShdR.style.MozOpacity=.5
			objShdB.style.MozOpacity=.5
		}
		if(objOver)
		{
			objOver.style.left=x+'px'
			objOver.style.top=y+'px'
			objOver.style.display="none"
		}
		if(objOverShadow)
		{
			objOverShadow.style.left=x+'px'
			objOverShadow.style.top=y+'px'
			objOverShadow.style.display="none"
		}
		objShdR.style.visibility="hidden"
		objShdR.style.left=x+'px'
		objShdR.style.top=y+'px'
		objShdB.style.visibility="hidden"
		objShdB.style.left=x+'px'
		objShdB.style.top=y+'px'
	}
	if(objPopCal.timeoutID3)
	{
		clearTimeout(objPopCal.timeoutID3)
		objPopCal.timeoutID3=null
	}
}

function PopCalIsObjectVisible(obj)
{
	var bVisible=((obj.style.display!='none')&&(obj.style.visibility!='hidden'))
	var objParent=obj.offsetParent
	while((objParent)&&(objParent.tagName!="BODY")&&(objParent.tagName!="HTML")&&(bVisible))
	{
		bVisible=((objParent.style.display!='none')&&(objParent.style.visibility!='hidden'))
		objParent=objParent.offsetParent
	}
	return(bVisible)
}

function PopCalConstructDate(d,m,y,format,l)
{
	var objPopCal=objPopCalList[l]
	var sTmp=format
	sTmp=sTmp.replace("dd","<e>")
	sTmp=sTmp.replace("d","<d>")
	sTmp=sTmp.replace("<e>",PopCalPad(d,2,"0","L"))
	sTmp=sTmp.replace("<d>",d)
	sTmp=sTmp.replace("mmmm","<l>")
	sTmp=sTmp.replace("mmm","<s>")
	sTmp=sTmp.replace("mm","<n>")
	sTmp=sTmp.replace("m","<m>")
	sTmp=sTmp.replace("yyyy",PopCalPad(y,4,"0","L"))
	sTmp=sTmp.replace("yy",PopCalPad(y,4,"0","L").substr(2))
	sTmp=sTmp.replace("<m>",m+1)
	sTmp=sTmp.replace("<n>",PopCalPad(m+1,2,"0","L"))
	sTmp=sTmp.replace("<s>",objPopCal.monthNameShort[m])
	sTmp=sTmp.replace("<l>",objPopCal.monthName[m])
	return(sTmp)
}

function PopCalCloseCalendar(l)
{
	var objPopCal=objPopCalList[l]
	clearInterval(objPopCal.intervalID1)
	clearTimeout(objPopCal.timeoutID1)
	clearInterval(objPopCal.intervalID2)
	clearTimeout(objPopCal.timeoutID2)
	PopCalHideCalendar(l)
	if(!objPopCal)
	{
		objPopCal=null
		return(false)
	}
	objPopCal.ctl().value=PopCalConstructDate(objPopCal.dateSelected,objPopCal.monthSelected,objPopCal.yearSelected,objPopCal.dateFormat,l)
	if(objPopCal.commandExecute)
	{
		eval(objPopCal.commandExecute)
	}
	else
	{
		PopCalSetFocus(objPopCal.ctl())
	}
}

function PopCalClickDocumentBody(l)
{
	var objPopCal=objPopCalList[l]
	if(objPopCal.ie)
	{
		if(event.keyCode==82)
		{
			var obj=objPopCal.ctl()
			if(obj)
			{
				if((obj.CalendarLeft)&&(obj.CalendarTop))
				{
					PopCalMoveDefault(l)
					PopCalDrop(l)
					if(document.body)
					{
						PopCalSetFocus(document.body)
					}
				}
			}
		}
	}
	PopCalGetById("popupSuperHighLight"+objPopCal.id).style.borderColor="#a0a0a0"
	objPopCal.PopCalDragClose=false
	if(!objPopCal.bShow)
	{
		PopCalHideCalendar(l)
	}
	if(!objPopCal)
	{
		objPopCal=null
		return(false)
	}
	objPopCal.bShow=false
}

function PopCalStartDecMonth(l)
{
	var objPopCal=objPopCalList[l]
	PopCalDownMonth(l)
	PopCalDownYear(l)
	clearInterval(objPopCal.intervalID1)
	clearTimeout(objPopCal.timeoutID1)
	clearInterval(objPopCal.intervalID2)
	clearTimeout(objPopCal.timeoutID2)
	objPopCal.intervalID1=setInterval("PopCalDecMonth("+l+")",80)
}

function PopCalStartIncMonth(l)
{
	var objPopCal=objPopCalList[l]
	PopCalDownMonth(l)
	PopCalDownYear(l)
	clearInterval(objPopCal.intervalID1)
	clearTimeout(objPopCal.timeoutID1)
	clearInterval(objPopCal.intervalID2)
	clearTimeout(objPopCal.timeoutID2)
	objPopCal.intervalID1=setInterval("PopCalIncMonth("+l+")",80)
}

function PopCalIncMonth(l)
{
	var objPopCal=objPopCalList[l]
	objPopCal.monthSelected++
	if(objPopCal.monthSelected>11)
	{
		objPopCal.monthSelected=0
		objPopCal.yearSelected++
	}
	if((objPopCal.yearSelected>objPopCal.yearUpTo)||((objPopCal.yearSelected==objPopCal.yearUpTo)&&(objPopCal.monthSelected>objPopCal.monthUpTo)))
	{
		PopCalDecMonth(l)
	}
	else
	{
		PopCalConstructCalendar(l)
		if(objPopCal.lr==0)
		{
			PopCalMoveDefaultPos(l)
			PopCalMoveShadow(l)
		}
	}
}

function PopCalDecMonth(l)
{
	var objPopCal=objPopCalList[l]
	objPopCal.monthSelected--
	if(objPopCal.monthSelected<0)
	{
		objPopCal.monthSelected=11
		objPopCal.yearSelected--
	}
	if((objPopCal.yearSelected<objPopCal.yearFrom)||((objPopCal.yearSelected==objPopCal.yearFrom)&&(objPopCal.monthSelected<objPopCal.monthFrom)))
	{
		PopCalIncMonth(l)
	}
	else
	{
		PopCalConstructCalendar(l)
		if(objPopCal.lr==0)
		{
			PopCalMoveDefaultPos(l)
			PopCalMoveShadow(l)
		}
	}
}

function PopCalConstructMonth(l)
{
	var objPopCal=objPopCalList[l]
	PopCalDownYear(l)
	if(!objPopCal.monthConstructed)
	{
		var beginMonth=0
		var endMonth=11
		objPopCal.countMonths=0
		if(objPopCal.yearSelected==objPopCal.yearFrom)
		{
			beginMonth=objPopCal.monthFrom
		}
		if(objPopCal.yearSelected==objPopCal.yearUpTo)
		{
			endMonth=objPopCal.monthUpTo
		}
		var sHTML=""
		for(var i=beginMonth;i<=endMonth;i++)
		{
			objPopCal.countMonths++
			var sName=objPopCal.monthName[i]
			if(i==objPopCal.monthSelected)
			{
				sName="<b>"+sName+"</b>"
			}
			sHTML+="<tr><td id='popupSuperMonth"+i+"' class='OptionStyle' onmouseover='objPopCalList["+l+"].bShow=true;this.className=\"OptionOverStyle\"' onmouseout='objPopCalList["+l+"].bShow=false;this.className=\"OptionOutStyle\"' style='cursor:default;' onclick='objPopCalList["+l+"].bShow=false;objPopCalList["+l+"].monthConstructed=false;objPopCalList["+l+"].monthSelected="+i+";PopCalConstructCalendar("+l+");PopCalDownMonth("+l+");'>&nbsp;"+sName+"&nbsp;</td></tr>"
		}
		var PopCal=objPopCal.Object
		PopCal.popupSuperMonth.className=objPopCal.CssClass
		PopCal.popupSuperMonth.innerHTML="<table dir='"+((objPopCal.lr==1)?"ltr":"rtl")+"' width='100%' style='border:1px solid #a0a0a0;' class='ListStyle' cellspacing=0 onmouseover='clearTimeout(objPopCalList["+l+"].timeoutID1)' onmouseout='clearTimeout(objPopCalList["+l+"].timeoutID1);'>"+sHTML+"</table>"
		objPopCal.monthConstructed=true
	}
}

function PopCalUpMonth(l)
{
	var objPopCal=objPopCalList[l]
	var PopCal=objPopCal.Object
	var objSpanMonth=PopCalGetById("popupSuperSpanMonth"+PopCal.id)
	if((objPopCal.yearSelected==objPopCal.yearFrom)||(objPopCal.yearSelected==objPopCal.yearUpTo))
	{
		objPopCal.monthConstructed=false
	}
	else if(objPopCal.countMonths!=12)
	{
		objPopCal.monthConstructed=false
	}
	PopCalConstructMonth(l)
	PopCal.popupSuperMonth.style.display=""
	if(PopCal.popupSuperMonth.OverSelect) PopCal.popupSuperMonth.OverSelect.style.display=""
	var lTop=parseInt(PopCal.popupSuperCalendar.style.top,10)+objSpanMonth.offsetTop+objSpanMonth.offsetHeight+6
	var lLeft=parseInt(PopCal.popupSuperCalendar.style.left,10)+objSpanMonth.offsetLeft
	if(objPopCal.lr==0)
	{
		if(objPopCal.ie)
		{
			lLeft=parseInt(PopCal.popupSuperCalendar.style.left,10)+PopCal.popupSuperCalendar.offsetWidth-(objSpanMonth.offsetLeft+objSpanMonth.offsetWidth)
			lLeft-=8
		}
		else
		{
			var _tl=PopCalGetTopLeft(objSpanMonth)
			lLeft=_tl[1]
		}
	}
	else
	{
		lLeft+=((objPopCal.ie)?4:5)
	}
	PopCalSetPosition(PopCal.popupSuperMonth,lTop,lLeft,objSpanMonth.offsetHeight,objSpanMonth.offsetWidth)
	if((objPopCal.lr==0)&&(objPopCal.ie))
	{
		if(objSpanMonth.offsetWidth<PopCal.popupSuperMonth.offsetWidth)
		{
			lLeft-=(PopCal.popupSuperMonth.offsetWidth-objSpanMonth.offsetWidth)
		}
		PopCalSetPosition(PopCal.popupSuperMonth,lTop,lLeft)
	}
}

function PopCalDownMonth(l)
{
	var objPopCal=objPopCalList[l]
	var obj=objPopCal.Object.popupSuperMonth
	if(obj.style.display!="none")
	{
		if(!objPopCal.keepMonth)
		{
			clearInterval(objPopCal.intervalID1)
			clearTimeout(objPopCal.timeoutID1)
			clearInterval(objPopCal.intervalID2)
			clearTimeout(objPopCal.timeoutID2)
			obj.style.display="none"
			obj=obj.OverSelect
			if(obj)
			{
				obj.style.display="none"
				obj.style.height='1px'
				obj.style.width='1px'
			}
		}
	}
	objPopCal.keepMonth=false
}

function PopCalWheelYear(l)
{
	var objPopCal=objPopCalList[l]
	if(PopCalIsObjectVisible(objPopCal.Object.popupSuperYear))
	{
		if(event)
		{
			if(event.wheelDelta>=120)
			{
				for(var i=0;i<3;i++)
				{
					PopCalDecYear(l)
				}
			}
			else if(event.wheelDelta<=-120)
			{
				for(var i=0;i<3;i++)
				{
					PopCalIncYear(l)
				}
			}
			event.returnValue=false
		}
	}
}

function PopCalIncYear(l)
{
	var objPopCal=objPopCalList[l]
	if((objPopCal.nStartingYear+(objPopCal.HalfYearList*2+1))<=objPopCal.yearUpTo)
	{
		var PopCal=objPopCal.Object
		for(var i=0;i<(objPopCal.HalfYearList*2+1);i++)
		{
			var newYear=(i+objPopCal.nStartingYear)+1
			var txtYear
			if(newYear==objPopCal.yearSelected)
			{
				txtYear="&nbsp;<b>"+newYear+"</b>&nbsp;"
			}
			else
			{
				txtYear="&nbsp;"+newYear+"&nbsp;"
			}
			PopCal.popupSuperYearList[i].innerHTML=txtYear
		}
		objPopCal.nStartingYear++
	}
	objPopCal.bShow=true
}

function PopCalDecYear(l)
{
	var objPopCal=objPopCalList[l]
	if(objPopCal.nStartingYear-1>=objPopCal.yearFrom)
	{
		var PopCal=objPopCal.Object
		for(var i=0;i<(objPopCal.HalfYearList*2+1);i++)
		{
			var newYear=(i+objPopCal.nStartingYear)-1
			var txtYear
			if(newYear==objPopCal.yearSelected)
			{
				txtYear="&nbsp;<b>"+ newYear+"</b>&nbsp;"
			}
			else
			{
				txtYear="&nbsp;"+newYear+"&nbsp;"
			}
			PopCal.popupSuperYearList[i].innerHTML=txtYear
		}
		objPopCal.nStartingYear--
	}
	objPopCal.bShow=true
}

function PopCalSelectYear(nYear,l)
{
	var objPopCal=objPopCalList[l]
	objPopCal.yearSelected=nYear+objPopCal.nStartingYear
	if((objPopCal.yearSelected==objPopCal.yearFrom)&&(objPopCal.monthSelected<objPopCal.monthFrom))
	{
		objPopCal.monthSelected=objPopCal.monthFrom
	}
	else if((objPopCal.yearSelected==objPopCal.yearUpTo)&&(objPopCal.monthSelected>objPopCal.monthUpTo))
	{
		objPopCal.monthSelected=objPopCal.monthUpTo
	}
	objPopCal.yearConstructed=false
	PopCalConstructCalendar(l)
	PopCalDownYear(l)
}

function PopCalConstructYear(l)
{
	var objPopCal=objPopCalList[l]
	PopCalDownMonth(l)
	var sHTML=""
	var longList=true
	if(!objPopCal.yearConstructed)
	{
		var beginYear=objPopCal.yearSelected-objPopCal.HalfYearList
		var endYear=objPopCal.yearSelected+objPopCal.HalfYearList
		if((objPopCal.yearUpTo-objPopCal.yearFrom+1)<=(objPopCal.HalfYearList*2+1))
		{
			beginYear=objPopCal.yearFrom
			endYear=objPopCal.yearUpTo
			longList=false
		}
		else if(beginYear<objPopCal.yearFrom)
		{
			beginYear=objPopCal.yearFrom
			endYear=beginYear+objPopCal.HalfYearList*2
		}
		else if(endYear>objPopCal.yearUpTo)
		{
			endYear=objPopCal.yearUpTo
			beginYear=endYear-(objPopCal.HalfYearList*2)
		}
		objPopCal.nStartingYear=beginYear
		if(longList)
		{
			sHTML+="<tr><td align='center' class='OptionStyle' onmouseover='objPopCalList["+l+"].bShow=true;this.className=\"OptionOverStyle\"' onmouseout='objPopCalList["+l+"].bShow=false;clearInterval(objPopCalList["+l+"].intervalID1);this.className=\"OptionOutStyle\"' style='cursor:default;border-bottom:1px #a0a0a0 solid;' onmousedown='clearInterval(objPopCalList["+l+"].intervalID1);objPopCalList["+l+"].intervalID1=setInterval(\"PopCalDecYear("+l+")\",10)' onmouseup='clearInterval(objPopCalList["+l+"].intervalID1)'><img id='popupSuperUpYear' ondrag='return(false)' src='"+objPopCal.imgDir+"up.gif' border=0 /></td></tr>"
		}
		var j=0
		for(var i=(beginYear);i<=(endYear);i++)
		{
			var sName=i
			if(i==objPopCal.yearSelected)
			{
				sName="<b>"+sName+"</b>"
			}
			sHTML+="<tr><td id='popupSuperYear"+j+"' class='OptionStyle' align='center' onmouseover='objPopCalList["+l+"].bShow=true;this.className=\"OptionOverStyle\"' onmouseout='objPopCalList["+l+"].bShow=false;this.className=\"OptionOutStyle\"' style='cursor:default;' onclick='objPopCalList["+l+"].bShow=false;PopCalSelectYear("+j+","+l+");'>&nbsp;"+sName+"&nbsp;</td></tr>"
			j++
		}
		if(longList)
		{
			sHTML+="<tr><td align='center' class='OptionStyle' onmouseover='objPopCalList["+l+"].bShow=true;this.className=\"OptionOverStyle\"' onmouseout='objPopCalList["+l+"].bShow=false;clearInterval(objPopCalList["+l+"].intervalID2);this.className=\"OptionOutStyle\"' style='cursor:default;border-top:1px #a0a0a0 solid;' onmousedown='clearInterval(objPopCalList["+l+"].intervalID2);objPopCalList["+l+"].intervalID2=setInterval(\"PopCalIncYear("+l+")\",10)' onmouseup='clearInterval(objPopCalList["+l+"].intervalID2)'><img id='popupSuperDownYear' ondrag='return(false)' src='"+objPopCal.imgDir+"down.gif' border=0 /></td></tr>"
		}
		var PopCal=objPopCal.Object
		PopCal.popupSuperYear.className=objPopCal.CssClass
		PopCal.popupSuperYear.innerHTML="<table dir='"+((objPopCal.lr==1)?"ltr":"rtl")+"' width='100%' style='border:1px solid #a0a0a0;' class='ListStyle' onmouseover='clearTimeout(objPopCalList["+l+"].timeoutID2)' onmouseout='clearTimeout(objPopCalList["+l+"].timeoutID2);' cellspacing=0>"	+ sHTML	+ "</table>"
		PopCal.popupSuperYearList=[]
		for(var i=0;i<j;i++)
		{
			PopCal.popupSuperYearList[i]=PopCalGetById("popupSuperYear"+i)
		}
		objPopCal.yearConstructed=true
	}
}
function PopCalDownYear(l)
{
	var objPopCal=objPopCalList[l]
	obj=objPopCal.Object.popupSuperYear
	if(obj.style.display!="none")
	{
		if(!objPopCal.keepYear)
		{
			clearInterval(objPopCal.intervalID1)
			clearTimeout(objPopCal.timeoutID1)
			clearInterval(objPopCal.intervalID2)
			clearTimeout(objPopCal.timeoutID2)
			PopCalYearDown=true
			obj.style.display="none"
			obj=obj.OverSelect
			if(obj)
			{
				obj.style.display="none"
				obj.style.height='1px'
				obj.style.width='1px'
			}
		}
	}
	objPopCal.keepYear=false
}

function PopCalUpYear(l)
{
	var objPopCal=objPopCalList[l]
	var PopCal=objPopCal.Object
	var objSpanYear=PopCalGetById("popupSuperSpanYear"+PopCal.id)
	PopCalConstructYear(l)
	PopCal.popupSuperYear.style.display=""
	if(PopCal.popupSuperYear.OverSelect) PopCal.popupSuperYear.OverSelect.style.display=""
	var lTop=parseInt(PopCal.popupSuperCalendar.style.top,10)+objSpanYear.offsetTop+objSpanYear.offsetHeight+6
	var lLeft=parseInt(PopCal.popupSuperCalendar.style.left,10)+objSpanYear.offsetLeft
	if(objPopCal.lr==0)
	{
		if(objPopCal.ie)
		{
			lLeft=parseInt(PopCal.popupSuperCalendar.style.left,10)+PopCal.popupSuperCalendar.offsetWidth-(objSpanYear.offsetLeft+objSpanYear.offsetWidth)
			lLeft-=8
		}
		else
		{
			var _tl=PopCalGetTopLeft(objSpanYear)
			lLeft=_tl[1]
		}
	}
	else
	{
		lLeft+=((objPopCal.ie)?4:5)
	}
	PopCalSetPosition(PopCal.popupSuperYear,lTop,lLeft,objSpanYear.offsetHeight,objSpanYear.offsetWidth)
}

function PopCalGetWeekNumber1(_d,l)
{
	var objPopCal=objPopCalList[l]
	var PopCal=objPopCal.Object
	return(PopCalWeekNbr(_d,PopCal.startAt,PopCal.weekNumberFormula))
}

function PopCalWeekNbr(_n,_s,_c)
{
	var _y=_n.getFullYear()
	var _m=_n.getMonth()
	var _d=_n.getDate()
	_w0=PopCalGetWeekNumber(_y,_m,_d,_s,_c)
	if(_c==1)
	{
		if(_w0>52)
		{
			var _d0=PopCalSetDays(_n,7)
			var _d1=new Date(_y+1,0,1)
			if(_d0>_d1) _w0=1
		}
	}
	else
	{
		var _w1=0
		if(_w0==0)
		{
			_w1=PopCalGetWeekNumber(_y-1,11,31,_s,_c);
		}
		else if(_w0>52)
		{
			_w1=PopCalGetWeekNumber(_y+1,0,1,_s,_c)
		}
		if(_w1>0) _w0=_w1
	}
	return(_w0)
}

function PopCalConstructCalendar(l)
{
	var objPopCal=objPopCalList[l]
	var PopCal=objPopCal.Object
	var aNumDays=[31,0,31,30,31,30,31,31,30,31,30,31]
	var startDate=new Date(objPopCal.yearSelected,objPopCal.monthSelected,1)
	var endDate
	var numDaysInMonth
	var notSelect
	var selectWeekends=PopCal.selectWeekend
	var selectHolidays=PopCal.selectHoliday
	var _DayOfWeek=0
	var sAlign=((objPopCal.lr==1)?"right":"left")
	var sPad="&nbsp;"
	//Martes de Carnaval y Viernes Santo para el año actual
	if((PopCal.addCarnival>=1)||(PopCal.addGoodFriday>=1))
	{
		var dtDomingoPascuas=PopCalDomingoPascuas(objPopCal.yearSelected)
		if(PopCal.addCarnival>=1)
		{
			var dtDate=PopCalSetDays(dtDomingoPascuas,-47)
			if(PopCal.addCarnival==1)
			{
				PopCalAddHoliday(dtDate.getDate(),dtDate.getMonth()+1,dtDate.getFullYear(),objPopCal.CarnivalString,l)
			}
			else
			{
				PopCalAddSpecialDay(dtDate.getDate(),dtDate.getMonth()+1,dtDate.getFullYear(),objPopCal.CarnivalString,l)
			}
		}
		if(PopCal.addGoodFriday>=1)
		{
			var dtDate=PopCalSetDays(dtDomingoPascuas,-2)
			if(PopCal.addGoodFriday==1)
			{
				PopCalAddHoliday(dtDate.getDate(),dtDate.getMonth()+1,dtDate.getFullYear(),objPopCal.GoodFridayString,l)
			}
			else
			{
				PopCalAddSpecialDay(dtDate.getDate(),dtDate.getMonth()+1,dtDate.getFullYear(),objPopCal.GoodFridayString,l)
			}
		}
	}
	if(PopCal.showHolidays==0)
	{
		selectHolidays=1
	}
	if(objPopCal.monthSelected==1)
	{
		endDate=PopCalSetDays(new Date(objPopCal.yearSelected,2,1),-1)
		numDaysInMonth=endDate.getDate()
	}
	else
	{
		numDaysInMonth=aNumDays[objPopCal.monthSelected]
	}
	var dayPointer=startDate.getDay()-PopCal.startAt
	if(dayPointer<0)
	{
		dayPointer=(7-PopCal.startAt)+startDate.getDay()
	}
	var sHTML="<table border=0 cellpadding=0 cellspacing=0 class='BodyStyle'><tr>"
	if(PopCal.showWeekNumber==1)
	{
		sHTML+="<td class='HeaderStyle' align='"+sAlign+"' style='border-"+sAlign+":1px solid #a0a0a0!important;'>&nbsp;</td>"
	}
	for(var i=PopCal.startAt;i<7;i++)
	{
		sHTML+="<td class='HeaderStyle' align='"+sAlign+"'>"+sPad+objPopCal.dayName[i]+"</td>"
	}
	for(var i=0;i<PopCal.startAt;i++)
	{
		sHTML+="<td class='HeaderStyle' align='"+sAlign+"'>"+sPad+objPopCal.dayName[i]+"</td>"
	}
	sHTML+="</tr><tr>"
	if(PopCal.showWeekNumber==1)
	{
		sHTML+="<td class='DateStyle WeekNumberStyle' align='"+sAlign+"' style='border-"+sAlign+":1px solid #a0a0a0!important;'>"+PopCalWeekNbr(startDate,PopCal.startAt,PopCal.weekNumberFormula)+sPad+"</td>"
	}
	var _date=new Date(objPopCal.yearSelected,objPopCal.monthSelected,1)
	var _EndDate=new Date(objPopCal.yearSelected,objPopCal.monthSelected,numDaysInMonth)
	if(PopCal.showDaysOutMonth!=1)
	{
		for(var i=1;i<=dayPointer;i++)
		{
			sHTML+="<td class='DateStyle'>&nbsp;</td>"
		}
	}
	else
	{
		_date=PopCalSetDays(_date,-dayPointer)
		dayPointer=0
		_EndDate=PopCalSetDays(_EndDate,1)
		while(_EndDate.getDay()!=PopCal.startAt)
		{
			_EndDate=PopCalSetDays(_EndDate,1)
		}
		_EndDate=PopCalSetDays(_EndDate,-1)
	}
	do
	{
		dayPointer++
		sHTML+="<td align="+sAlign+" class='DateStyle'>"
		var sStyle=""
		if(objPopCal.monthSelected!=_date.getMonth())
		{
			sStyle+=" DaysOutOfMonthStyle"
		}
		if((_date.getDate()==objPopCal.odateSelected)&&(_date.getMonth()==objPopCal.omonthSelected)&&(_date.getFullYear()==objPopCal.oyearSelected))
		{
			sStyle+=" SelectedDateStyle"
		}
		notSelect=false
		var bHoliday=false
		var bSpecial=false
		var _IsDate=false
		var sHint=""
		var _reason=""
		for(var k=0;k<objPopCal.HolidaysCounter;k++)
		{
			_IsDate=false
			if(objPopCal.Holidays[k].type=="Type 1")
			{
				_IsDate=PopCalValidateType1(_date,objPopCal.Holidays[k])
			}
			else if(objPopCal.Holidays[k].type=="Type 2")
			{
				_IsDate=PopCalValidateType2(_date,objPopCal.Holidays[k])
			}
			else if(objPopCal.Holidays[k].type=="Type 3")
			{
				_IsDate=PopCalValidateType3(_date,objPopCal.Holidays[k])
			}
			if(_IsDate)
			{
				if((PopCal.showHolidays==1)&&(objPopCal.Holidays[k].tipo==1))
				{
					bHoliday=true
					sHint+=(sHint==""?"":(objPopCal.ie?"\n":", "))+objPopCal.Holidays[k].desc
				}
				else if((PopCal.showSpecialDay==1)&&(objPopCal.Holidays[k].tipo==0))
				{
					bSpecial=true
					sHint+=(sHint==""?"":(objPopCal.ie?"\n":", "))+objPopCal.Holidays[k].desc
				}
				if((selectHolidays!=1)&&(objPopCal.Holidays[k].tipo==1))
				{
					notSelect=true
					if(_reason!="") _reason+=","
					_reason+="Holiday"
				}
			}
		}
		var regexp=/\"/g
		sHint=sHint.replace(regexp,"&quot;")
		if(PopCalDateProcess(_date)<PopCalDateFrom(l))
		{
			notSelect=true
			if(_reason!="") _reason+=","
			_reason+="RangeFrom"
		}
		if(PopCalDateUpTo(l)<PopCalDateProcess(_date))
		{
			notSelect=true
			if(_reason!="") _reason+=","
			_reason+="RangeTo"
		}
		if(selectWeekends!=1)
		{
			if(PopCal.weekend.indexOf(_date.getDay().toString())!=-1)
			{
				notSelect=true
				if(_reason!="") _reason+=","
				_reason+="Weekend"
			}
		}
		if(PopCalDateNow(l)==PopCalDateProcess(_date))
		{
			sStyle+=" CurrentDateStyle"
		}
		var _Style=''
		var _S1=''
		var _S2=''
		var _S3=''
		var _S4=''
		if((PopCal.weekend.indexOf(_date.getDay().toString())!=-1)&&(PopCal.showWeekend==1))
		{
			if(objPopCal.ClientScriptWeekendDateStyle!='')
			{
				if(typeof(eval("window."+objPopCal.ClientScriptWeekendDateStyle))=="function")
				{
					_S1=eval("window."+objPopCal.ClientScriptWeekendDateStyle+"(_date,objPopCal,sHint)")
					if(typeof(_S1)!="string") _S1=''
				}
			}
			if(_S1!='')
			{
				_Style=_S1
			}
			else
			{
				sStyle+=" WeekendStyle"
			}
		}
		if(bSpecial)
		{
			if(objPopCal.ClientScriptSpecialDateStyle!='')
			{
				if(typeof(eval("window."+objPopCal.ClientScriptSpecialDateStyle))=="function")
				{
					_S2=eval("window."+objPopCal.ClientScriptSpecialDateStyle+"(_date,objPopCal,sHint)")
					if(typeof(_S2)!="string") _S2=''
				}
			}
			if(_S2!='')
			{
				_Style=_S2
			}
			else
			{
				sStyle+=" SpecialDayStyle"
			}
		}
		if(bHoliday)
		{
			if(objPopCal.ClientScriptHolidayDateStyle!='')
			{
				if(typeof(eval("window."+objPopCal.ClientScriptHolidayDateStyle))=="function")
				{
					_S3=eval("window."+objPopCal.ClientScriptHolidayDateStyle+"(_date,objPopCal,sHint)")
					if(typeof(_S3)!="string") _S3=''
				}
			}
			if(_S3!='')
			{
				_Style=_S3
			}
			else
			{
				sStyle+=" HolidayStyle"
			}
		}
		if(notSelect)
		{
			if(objPopCal.ClientScriptDisabledDateStyle!='')
			{
				if(typeof(eval("window."+objPopCal.ClientScriptDisabledDateStyle))=="function")
				{
					_S4=eval("window."+objPopCal.ClientScriptDisabledDateStyle+"(_date,objPopCal,sHint,_reason)")
					if(typeof(_S4)!="string") _S4=''
				}
			}
			if(_S4!='')
			{
				_Style=_S4
			}
			else
			{
				sStyle+=" DisableDateStyle"
			}
		}
		if((_S4!='')&&(_S3!=''))
		{
			sStyle+=" HolidayStyle"
		}
		if(((_S4+_S3)!='')&&(_S2!=''))
		{
			sStyle+=" SpecialDayStyle"
		}
		if(((_S4+_S3+_S2)!='')&&(_S1!=''))
		{
			sStyle+=" WeekendStyle"
		}
		if(_Style.indexOf(":")!=-1)
		{
			_Style=" Style='"+_Style+"'"
		}
		else
		{
			if(_Style!='') sStyle+=(" "+_Style)
			_Style=""
		}
		if(notSelect)
		{
			sHTML+="<span title=\""+sHint+"\" class='"+sStyle+"' "+_Style+">&nbsp;"+_date.getDate()+"&nbsp;</span>"
		}
		else
		{
			var _formatMsg=''
			if(objPopCal.lr==0)
			{
				for(var i=objPopCal.dateFormat.length-1;i>=0;i--)
				{
					_formatMsg+=objPopCal.dateFormat.substr(i,1)
				}
			}
			else
			{
				_formatMsg=objPopCal.dateFormat
			}
			var dateMessage="onmouseover='window.status=\""+objPopCal.selectDateMessage.replace("[Date]",PopCalConstructDate(_date.getDate(),_date.getMonth(),_date.getFullYear(),_formatMsg,l))+"\";this.className+=\" DayOverStyle\";' onmouseout='window.status=\"\";this.className=this.getAttribute(\"CSS\");' "
			sHTML+="<span "+dateMessage+" title=\""+sHint+"\" class='"+sStyle+"' CSS='"+sStyle+"' "+_Style+" onclick='objPopCalList["+l+"].dateSelected="+_date.getDate()+";objPopCalList["+l+"].monthSelected="+_date.getMonth()+";objPopCalList["+l+"].yearSelected="+_date.getFullYear()+";PopCalCloseCalendar("+l+");'>&nbsp;"+_date.getDate()+"&nbsp;</span>"
		}
		sHTML+="</td>"
		if((PopCal.showDaysOutMonth!=1)||(PopCalDateProcess(_EndDate)>PopCalDateProcess(PopCalSetDays(_date,1))))
		{
			if((dayPointer+PopCal.startAt)%7==PopCal.startAt)
			{
				sHTML+="</tr><tr>"
				if((PopCal.showWeekNumber==1)&&(_date.getDate()<numDaysInMonth))
				{
					sHTML+="<td class='DateStyle WeekNumberStyle' align='"+sAlign+"' style='border-"+sAlign+":1px solid #a0a0a0!important;'>"+(PopCalWeekNbr(new Date(_date.getFullYear(),_date.getMonth(),_date.getDate()+1),PopCal.startAt,PopCal.weekNumberFormula))+sPad+"</td>"
				}
			}
		}
		_date=PopCalSetDays(_date,1)
	}
	while(PopCalDateProcess(_date)<=PopCalDateProcess(_EndDate))
	if(PopCal.showDaysOutMonth!=1)
	{
		while((dayPointer+PopCal.startAt)%7!=PopCal.startAt)
		{
			sHTML+="<td class='DateStyle'>&nbsp;</td>"
			++dayPointer
		}
	}
	sHTML+="</tr></table>"
	if(PopCal.addGoodFriday>=1)
	{
		objPopCal.Holidays.length=--objPopCal.HolidaysCounter
	}
	if(PopCal.addCarnival>=1)
	{
		objPopCal.Holidays.length=--objPopCal.HolidaysCounter
	}
	PopCalGetById("popupSuperContent"+PopCal.id).innerHTML=sHTML
	if(PopCalGetById("popupSuperCaption"+PopCal.id).dropDown)
	{
		PopCalGetById("popupSuperSpanMonth"+PopCal.id).innerHTML="&nbsp;"+objPopCal.monthName[objPopCal.monthSelected]+"&nbsp;<img id='popupSuperChangeMonth"+PopCal.id+"' ondrag='return(false)' src='"+objPopCal.imgDir+"drop1.gif' width='12' height='10' border='0' />"
		PopCalGetById("popupSuperSpanYear"+PopCal.id).innerHTML="&nbsp;"+objPopCal.yearSelected+"&nbsp;<img id='popupSuperChangeYear"+PopCal.id+"' ondrag='return(false)' src='"+objPopCal.imgDir+"drop1.gif' width='12' height='10' border='0' />"
	}
	else
	{
		PopCalGetById("popupSuperMoveCalendar"+PopCal.id).innerHTML=objPopCal.monthName[objPopCal.monthSelected]+"&nbsp;"+objPopCal.yearSelected
	}
	PopCalMoveShadow(l)
}

function PopCalMoveShadow(l)
{
	var objPopCal=objPopCalList[l]
	var PopCal=objPopCal.Object
	PopCalSetPosition(PopCal.popupSuperCalendar)
	if(PopCal.shadow==1)
	{
		var objCal=PopCal.popupSuperCalendar
		var obj=PopCal.popupSuperShadowRight
		obj.style.height=(objCal.offsetHeight-10)+'px'
		obj.style.top=(parseInt(objCal.style.top,10)+10)+'px'
		if(objPopCal.lr==1)
		{
			obj.style.left=(objCal.offsetLeft+objCal.offsetWidth)+'px'
		}
		else
		{
			obj.style.left=(objCal.offsetLeft-10)+'px'
		}
		obj=PopCal.popupSuperShadowBottom
		obj.style.width=objCal.offsetWidth+'px'
		obj.style.top=(parseInt(objCal.style.top,10)+objCal.offsetHeight)+'px'
		if(objPopCal.lr==1)
		{
			obj.style.left=((objCal.offsetLeft+objCal.offsetWidth+10)-objCal.offsetWidth)+'px'
		}
		else
		{
			obj.style.left=(objCal.offsetLeft-10)+'px'
		}
	}
}

function PopCalDateProcess(_date)
{
	return(PopCalPad(_date.getFullYear(),4,"0","L")+PopCalPad(_date.getMonth(),2,"0","L")+PopCalPad(_date.getDate(),2,"0","L"))
}

function PopCalDateNow(l)
{
	var objPopCal=objPopCalList[l]
	return(PopCalPad(objPopCal.yearNow,4,"0","L")+PopCalPad(objPopCal.monthNow,2,"0","L")+PopCalPad(objPopCal.dateNow,2,"0","L"))
}

function PopCalDateSelect(l)
{
	var objPopCal=objPopCalList[l]
	return(PopCalPad(objPopCal.yearSelected,4,"0","L")+PopCalPad(objPopCal.monthSelected,2,"0","L")+PopCalPad(objPopCal.dateSelected,2,"0","L"))
}

function PopCalDateFrom(l)
{
	var objPopCal=objPopCalList[l]
	return(PopCalPad(objPopCal.yearFrom,4,"0","L")+PopCalPad(objPopCal.monthFrom,2,"0","L")+PopCalPad(objPopCal.dateFrom,2,"0","L"))
}

function PopCalDateUpTo(l)
{
	var objPopCal=objPopCalList[l]
	return(PopCalPad(objPopCal.yearUpTo,4,"0","L")+PopCalPad(objPopCal.monthUpTo,2,"0","L")+PopCalPad(objPopCal.dateUpTo,2,"0","L"))
}

function PopCalGetSeparator(dateFormat)
{
	var Separator=' /-.'
	for(var i=0;i<Separator.length;i++)
	{
		var formatChar=Separator.substr(i,1)
		if(dateFormat.split(formatChar).length==3)
		{
			return(formatChar)
		}
	}
	return(null)
}

function PopCalCenturyOn(dateFormat)
{
	var formatChar=PopCalGetSeparator(dateFormat)
	if(formatChar)
	{
		var aFormat=dateFormat.toLowerCase().split(formatChar)
		for(var i=0;i<3;i++)
		{
			if(aFormat[i]=="yyyy")
			{
				return(true)
			}
		}
	}
	return(false)
}

function PopCalSetDMY(dateValue,dateFormat,l)
{
	var objPopCal=objPopCalList[l]
	var PopCal=objPopCal.Object
	var tokensChanged=0
	var formatChar=PopCalGetSeparator(dateFormat)
	var _p=""
	objPopCal.oDate=null
	objPopCal.oMonth=null
	objPopCal.oYear=null
	if(formatChar)
	{
		if(formatChar==".")
		{
			if(dateValue.indexOf("..")!=-1)
			{
				_p="."
			}
		}
		// use user's date
		var aFormat=dateFormat.toLowerCase().split(formatChar)
		var aData=dateValue.toLowerCase().replace("..",".").split(formatChar)
		for(var i=0;i<3;i++)
		{
			if((aFormat[i]=="d")||(aFormat[i]=="dd"))
			{
				objPopCal.oDate=parseInt(aData[i],10)
				tokensChanged++
			}
			else if((aFormat[i]=="m")||(aFormat[i]=="mm"))
			{
				if(((parseInt(aData[i],10)-1)>=0)&&((parseInt(aData[i],10)-1)<=11))
				{
					objPopCal.oMonth=parseInt(aData[i],10)-1
					tokensChanged++
				}
			}
			else if((aFormat[i]=="yy")||(aFormat[i]=="yyyy"))
			{
				objPopCal.oYear=parseInt(aData[i],10)
				if(objPopCal.oYear<=99)
				{
					tokensChanged++
					if(objPopCal.oYear<100)
					{
						if(objPopCal.oYear<PopCal.centuryLimit)
						{
							objPopCal.oYear+=100
						}
						objPopCal.oYear+=1900
					}
				}
				else if(objPopCal.oYear<=9999)
				{
					tokensChanged++
				}
			}
			else if(aFormat[i]=="mmm")
			{
				for(j=0;j<12;j++)
				{
					if(aData[i])
					{
						if(aData[i]+_p==objPopCal.monthNameShort[j].toLowerCase())
						{
							objPopCal.oMonth=j
							tokensChanged++
							break
						}
						else if(aData[i]==objPopCal.monthNameShort[j].toLowerCase())
						{
							objPopCal.oMonth=j
							tokensChanged++
							break
						}
					}
				}
			}
			else if(aFormat[i]=="mmmm")
			{
				for(j=0;j<12;j++)
				{
					if(aData[i])
					{
						if(aData[i]==objPopCal.monthName[j].toLowerCase())
						{
							objPopCal.oMonth=j
							tokensChanged++
							break
						}
					}
				}
			}
		}
	}
	return((tokensChanged==3)&&!isNaN(objPopCal.oDate)&&!isNaN(objPopCal.oMonth)&&!isNaN(objPopCal.oYear))
}

function PopCalGetDate(dateValue,dateFormat,l)
{
	var objPopCal=objPopCalList[l]
	if(PopCalIsToday(dateValue))
	{
		return(new Date(objPopCal.today.getFullYear(),objPopCal.today.getMonth(),objPopCal.today.getDate()))
	}
	else if(PopCalSetDMY(dateValue,dateFormat,l))
	{
		return(new Date(objPopCal.oYear,objPopCal.oMonth,objPopCal.oDate))
	}
	return(null)
}

function PopCalChangeCurrentMonth(l)
{
	var objPopCal=objPopCalList[l]
	if((PopCalDateFrom(l).substr(0,6)<=PopCalDateNow(l).substr(0,6))&&(PopCalDateNow(l).substr(0,6)<=PopCalDateUpTo(l).substr(0,6)))
	{
		objPopCal.monthSelected=objPopCal.monthNow
		objPopCal.yearSelected=objPopCal.yearNow
		objPopCal.yearConstructed=false
		objPopCal.monthConstructed=false
		PopCalConstructCalendar(l)
	}
}

function PopCalDomingoPascuas(y)
{
	var lnCentena=parseInt(y/100,10)
	var lnAux=(y+1)%19
	var lnNroAureo=lnAux+(19*parseInt((19-lnAux)/19,10))
	var lnDomingo=7+(1-y-parseInt(y/4,10)+lnCentena-parseInt(lnCentena/4,10))%7
	var lnEpactaJul=((11*lnNroAureo)-10)%30
	var lnCorrSolar=-(lnCentena-16)+parseInt((lnCentena-16)/4,10)
	var lnCorrLunar=parseInt((lnCentena-15-parseInt((lnCentena-17)/25,10))/3,10)
	var lnEpactaGreg=(30+lnEpactaJul+lnCorrSolar+lnCorrLunar)%30
	var lnDiasLunaP=24-lnEpactaGreg+(30*parseInt(lnEpactaGreg/24,10))
	var lnDiasLuna15=(27-lnEpactaGreg+(30*parseInt(lnEpactaGreg/24,10)))%7
	var lnDiasPascua=lnDiasLunaP+(7+lnDomingo-lnDiasLuna15)%7
	var dtFecIni=new Date(y,2,21)
	var dtFecPascua=PopCalSetDays(dtFecIni,lnDiasPascua)
	return(dtFecPascua)
}

function PopCalGetWeekNumber(_y,_m,_d,_s,_c)
{
	var _dw=(7+(new Date(_y,0,1).getDay())-_s)%7
	var _cd=Date.UTC(_y,_m,_d)
	var _fdy=Date.UTC(_y,0,1)
	var _week=Math.ceil(((_cd-_fdy)/86400000+_dw-6)/7)
	if((_c==1)||(_dw<=3))
	{
		_week++
	}
	return(_week)
}

function PopCalPad(s,l,c,X)
{
	var x=X
	var r=s.toString()
	if(r.length>=l) return(r.substr(0,l))
	if(c==null) c=' '
	do
	{
		if(X=='C')
		{
			if(x=='L') x='R'
			else x='L'
		}
		if(x=='L') r=c+r
		else if(x=='R') r=r+c
	}
	while(r.length<l)
	return(r)
}

function PopCalIsToday(_hoy)
{
	return((',now,today,hoy,heute,oggi,hoje,').indexOf(','+_hoy.toLowerCase()+',')!=-1)
}

function PopCalIsGoodFriday(_date)
{
	var _goodFriday=PopCalSetDays(PopCalDomingoPascuas(_date.getFullYear()),-2)
	return(_date.toString()==_goodFriday.toString())
}

function PopCalIsCarnival(_date)
{
	var _carnival=PopCalSetDays(PopCalDomingoPascuas(_date.getFullYear()),-47)
	return(_date.toString()==_carnival.toString())
}

function PopCalGetById(_Id)
{
	return(document.getElementById(_Id))
}

function PopCalGetTopLeft(_obj)
{
	var t=_obj.offsetTop+_obj.offsetHeight
	var l=_obj.offsetLeft
	var objParent=_obj.offsetParent
	if(_obj.style.position.toLowerCase()!='absolute')
	{
		if(document.body)
		{
			if(document.body.offsetTop!=0)
			{
				if(document.body.currentStyle)
				{
					if(document.body.currentStyle.marginTop) t+=parseInt(document.body.currentStyle.marginTop,10)
					if(document.body.currentStyle.marginLeft) l+=parseInt(document.body.currentStyle.marginLeft,10)
				}
			}
		}
	}
	while((objParent.tagName!="BODY")&&(objParent.tagName!="HTML"))
	{
		t+=objParent.offsetTop-objParent.scrollTop
		l+=objParent.offsetLeft-objParent.scrollLeft
		if(objParent.tagName=="DIV")
		{
			if(!isNaN(parseInt(objParent.style.borderTopWidth,10))) t+=parseInt(objParent.style.borderTopWidth,10)
			if(!isNaN(parseInt(objParent.style.borderLeftWidth,10))) l+=parseInt(objParent.style.borderLeftWidth,10)
		}
		objParent=objParent.offsetParent
	}
	return([t,l])
}

function PopCalSetDays(_Date,_Inc)
{
	return (new Date((new Date(_Date)).setHours(_Inc*24,0,0,0)))
}
