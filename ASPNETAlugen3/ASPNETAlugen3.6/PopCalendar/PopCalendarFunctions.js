// Desarrollado por: Ricaute Jiménez Sánchez (La Chorrera, Panamá)
// Componente: PopCalendarFunctions.js(2.5.9)
// Correo: ricaj0625@yahoo.com
// Última actualización: 9 de mayo de 2006
// Por favor mantenga estos créditos (Please, keep these credits)

if(typeof(__PopCalValidCalendarRanges)=='undefined')
{
	var __PopCalValidCalendarRanges=[]
	var PopCalendarFunctions={majorVersion:2,minorVersion:5.9}
}

var __PopCalTemporal=""
var __PopCalLastControlFocus=""
var __PopCalTimerSummary=null
var __PopCalCheckDependencies=false

function __PopCalValidateKey(o,e)
{
	__PopCalLastControlFocus=""
	if(o.value!="")
	{
		if(e.target)
		{
			if(e.which==13)
			{
				__PopCalSetBlur(o,e)
			}
		}
	}
}

function __PopCalSetFocus(o,e)
{
	o.valueOnFocus=o.value
}

function __PopCalSetBlur(o,e)
{
	o.TempValue=o.value
	__PopCalFormatControl(o)
	if(o.value=="")
	{
		o.value=o.TempValue
	}
	o.TempValue=null
	if(o.value!=o.valueOnFocus)
	{
		var _PopCal=eval(o.getAttribute("Calendar"))
		__PopCalSelectionChanged(o,_PopCal)
	}
}

function __PopCalValidateOnSubmit(val,args)
{
	var o=document.getElementById(val.controltovalidate)
	if(!o)
	{
		args.IsValid=true
		return
	}
	o.value=__PopCalValueTrim(o.value)
	if((__PopCalCheckDependencies)&&((o.value=="")||(o.blankfield)))
	{
		if(typeof(val._IsValid)!="undefined")
		{
			args.IsValid=val._IsValid
		}
		return
	}
	if(typeof(val._IsValid)!="undefined")
	{
		if(Math.abs(__PopCalGetTicks()-val._ticks)<1000)
		{
			args.IsValid=val._IsValid
			return
		}
	}
	val._IsValid=true
	val._ticks=__PopCalGetTicks()
	var _PopCal=eval(o.getAttribute("Calendar"))
	var _format=o.getAttribute("Format")
	var _textMessage=o.getAttribute("TextMessage")
	if(_textMessage==null) _textMessage=""
	o.TempValue=o.value
	if(!__PopCalFormatControl(o))
	{
		o.value=o.TempValue
		o.TempValue =null
		if (_PopCal.Object.clientValidator==0) return(true)
		var _invalidDate=o.getAttribute("InvalidDateMessage")
		if((_invalidDate=="")||(_invalidDate==null)) _invalidDate="Invalid Date"
		_invalidDate=__PopCalSetErrorMessage(val,_invalidDate,_textMessage)
		args.IsValid=false
		val._IsValid=false
		val._ticks=__PopCalGetTicks()
		__PopCalShowMessage(o.id,_invalidDate)
		return
	}
	o.TempValue =null
	if(_PopCal.Object.clientValidator==0) return(true)
	if((o.value=="")&&(o.getAttribute("Required")=="true"))
	{
		var _focus=null
		if(event)
		{
			if(event.srcElement)
			{
				if(event.srcElement.id==val.controltovalidate)
				{
					_focus=false
				}
			}
		}
		var _requiredDate=o.getAttribute("RequiredDateMessage")
		if((_requiredDate=="")||(_requiredDate==null)) _requiredDate="Date is Required"
		_requiredDate=__PopCalSetErrorMessage(val,_requiredDate,_textMessage)
		args.IsValid=false
		val._IsValid=false
		val._ticks=__PopCalGetTicks()
		__PopCalShowMessage(o.id,_requiredDate,_focus)
		return
	}
	else if((o.value!=""))
	{
		var __Holiday=false
		var _date=_PopCal.getDate(o.value,_format)
		if(_PopCal.Object.showHolidays=="1")
		{
			if(_PopCal.Object.selectHoliday=="0")
			{
				var _holiday=o.getAttribute("HolidayMessage")
				if((_holiday=="")||(_holiday==null))
				{
					_holiday="Disabled Holidays"
				}
				for(var k=0;k<_PopCal.HolidaysCounter;k++)
				{
					if(_PopCal.Holidays[k].tipo==1)
					{
						__Holiday=false
						if(_PopCal.Holidays[k].type=="Type 1")
						{
							__Holiday=PopCalValidateType1(_date,_PopCal.Holidays[k])
						}
						else if(_PopCal.Holidays[k].type=="Type 2")
						{
							__Holiday=PopCalValidateType2(_date,_PopCal.Holidays[k])
						}
						if(__Holiday)
						{
							_holiday=__PopCalSetErrorMessage(val,_holiday,_textMessage)
							args.IsValid=false
							val._IsValid=false
							val._ticks=__PopCalGetTicks()
							__PopCalShowMessage(o.id,_holiday)
							return
						}
					}
				}
				var _DomingoPascuas=PopCalDomingoPascuas(_date.getFullYear())
				var _HolidayDate=new Date()
				if(_PopCal.Object.addCarnival=="1")
				{
					_HolidayDate=PopCalSetDays(_DomingoPascuas,-47)
					if(_HolidayDate.toString()==_date.toString())
					{
						_holiday=__PopCalSetErrorMessage(val,_holiday,_textMessage)
						args.IsValid=false
						val._IsValid=false
						val._ticks=__PopCalGetTicks()
						__PopCalShowMessage(o.id,_holiday)
						return
					}
				}
				if(_PopCal.Object.addGoodFriday=="1")
				{
					_HolidayDate=PopCalSetDays(_DomingoPascuas,-2)
					if(_HolidayDate.toString()==_date.toString())
					{
						_holiday=__PopCalSetErrorMessage(val,_holiday,_textMessage)
						args.IsValid=false
						val._IsValid=false
						val._ticks=__PopCalGetTicks()
						__PopCalShowMessage(o.id,_holiday)
						return
					}
				}
			}
		}
		if(_PopCal.Object.selectWeekend=="0")
		{
			var _weekend=o.getAttribute("WeekendMessage")
			if((_weekend=="")||(_weekend==null)) _weekend="Disabled Weekends"
			_date=_PopCal.getDate(o.value,_format)
			if(_PopCal.Object.weekend.indexOf(_date.getDay().toString())!=-1)
			{
				_weekend=__PopCalSetErrorMessage(val,_weekend,_textMessage)
				args.IsValid=false
				val._IsValid=false
				val._ticks=__PopCalGetTicks()
				__PopCalShowMessage(o.id,_weekend)
				return
			}
		}
		args.IsValid=__PopCalValidateRanges(o)
		val._IsValid=args.IsValid
		val._ticks=__PopCalGetTicks()
		if(args.IsValid)
		{
			if(!__PopCalCheckDependencies)
			{
				__PopCalCheckDependencies=true
				__PopCalValidateDependencies(o)
				__PopCalCheckDependencies=false
			}
		}
	}
	__PopCalUpdateSummaryValidator()
	if((o.valueOnFocus!=null)&&(o.value!=o.valueOnFocus)&&(args.IsValid)&&(o.blankfield==null))
	{
		__PopCalSelectionChanged(o,_PopCal)
	}
}

function __PopCalValidateRanges(o)
{
	var _PopCal=eval(o.getAttribute("Calendar"))
	var _format=o.getAttribute("Format")
	var _ValidControl=document.getElementById(o.getAttribute("Calendar")+"_MessageError")
	var _Range=null
	var _value=_PopCal.formatDate(o.value,_format,"yyyy-mm-dd")
	var _textMessage=o.getAttribute("TextMessage")
	if(_textMessage==null) _textMessage=""
	for(var i=0;i<__PopCalValidCalendarRanges.length;i++)
	{
		_Range=__PopCalValidCalendarRanges[i]
		if(_Range.Control==o.id)
		{
			break
		}
		_Range=null
	}
	if(_Range)
	{
		__PopCalTemporal=","
		var _DateFrom=__PopCalGetFromYYYYMMDD(o)
		if(_DateFrom!="")
		{
			if(_value<_DateFrom)
			{
				var _OutRange=_Range.FromMessage
				if(_OutRange=="") _OutRange="Out of Range"
				_OutRange=__PopCalSetErrorMessage(_ValidControl,_OutRange,_textMessage)
				__PopCalShowMessage(o.id,_OutRange)
				return(false)
			}
		}
		__PopCalTemporal=","
		var _DateTo=__PopCalGetToYYYYMMDD(o)
		if(_DateTo!="")
		{
			if(_value>_DateTo)
			{
				var _OutRange=_Range.ToMessage
				if(_OutRange=="") _OutRange="Out of Range"
				_OutRange=__PopCalSetErrorMessage(_ValidControl,_OutRange,_textMessage)
				__PopCalShowMessage(o.id,_OutRange)
				return(false)
			}
		}
	}
	return(true)
}

function __PopCalValidateDependencies(o)
{
	for(var i=0;i<__PopCalValidCalendarRanges.length;i++)
	{
		var _Range=__PopCalValidCalendarRanges[i]
		var _Control=document.getElementById(_Range.Control)
		if(_Control)
		{
			if(_Control.getAttribute("Buffer")!="true")
			{
				if(_Control.id!=o.id)
				{
					if((_Range.FromRange=="C:"+o.id)||(_Range.ToRange=="C:"+o.id))
					{
						if(typeof(Page_Validators)!="undefined")
						{
							for (j=0;j<Page_Validators.length;j++)
							{
								var val=Page_Validators[j]
								if(val.clientvalidationfunction=='__PopCalValidateOnSubmit')
								{
									if(val.controltovalidate==_Control.id)
									{
										if(val.isvalid)
										{
											if(_Control.value!="") _Control.valueOnFocus=_Control.value
											ValidatorValidate(val)
										}
									}
								}
							}
						}
					}
				}
			}
		}
	}
}

function __PopCalSelectNone(_id)
{
	PopCalHideCalendar(_id,false)
	for(var i=0;i<__PopCalValidCalendarRanges.length;i++)
	{
		var _Range=__PopCalValidCalendarRanges[i]
		var o=document.getElementById(_Range.Control)
		if(o)
		{
			var _PopCal=eval(o.getAttribute("Calendar"))
			if(_PopCal)
			{
				if(_PopCal.id==_id)
				{
					if(o.value!="") o.valueOnFocus=o.value
					o.value=""
					if(o.getAttribute("Buffer")=="true")
					{
						__PopCalSelectionChanged(o,_PopCal)
					}
					else if((typeof(ValidatorOnChange)=="function")&&(o.fireEvent))
					{
						o.fireEvent("onchange")
					}
					else if(o.valueOnFocus!=null)
					{
						if(o.value!=o.valueOnFocus)
						{
							__PopCalSelectionChanged(o,_PopCal)
						}
					}
					break
				}
			}
		}
	}
}

function __PopCalShowCalendar(_o,_span)
{
	var o=document.getElementById(_o)
	if(!o) return
	var _PopCal=eval(o.getAttribute("Calendar"))
	var _format=o.getAttribute("Format")
	var _from=""
	var _to=""
	o.value=__PopCalValueTrim(o.value)
	o.oldValue=o.value
	for(var i=0;i<__PopCalValidCalendarRanges.length;i++)
	{
		var _Range=__PopCalValidCalendarRanges[i]
		if(_Range.Control==o.id)
		{
			__PopCalTemporal=","
			var _f = _format
			if(_f.indexOf("yyyy")==-1) _f=_f.replace("yy","yyyy")
			var _DateFrom=__PopCalGetFromYYYYMMDD(o)
			_from=_PopCal.formatDate(_DateFrom,"yyyy-mm-dd",_f)
			__PopCalTemporal=","
			var _DateTo=__PopCalGetToYYYYMMDD(o)
			_to=_PopCal.formatDate(_DateTo,"yyyy-mm-dd",_f)
			break
		}
	}
	__PopCalLastControlFocus=""
	_PopCal.ControlAlignLeft=null
	if(o.getAttribute("Buffer")=="true")
	{
		_PopCal.ControlAlignLeft=_span
	}
	_PopCal.show(o,_format,_from,_to,"__PopCalSelectDate('"+o.id+"')")
}

function __PopCalSelectDate(_o)
{
	var o=document.getElementById(_o)
	if(!o) return
	if(o.value!=o.oldValue)
	{
		var _PopCal=eval(o.getAttribute("Calendar"))
		if(_PopCal)
		{
			if((_PopCal.ie)&&(_PopCal.Object.clientValidator==1)&&(o.getAttribute("Buffer")!="true"))
			{
				if ((typeof(ValidatorOnChange)=="function")&&(o.fireEvent))
				{
					o.fireEvent("onchange")
				}
				else if(!__PopCalCheckDependencies)
				{
					__PopCalCheckDependencies=true
					__PopCalValidateDependencies(o)
					__PopCalCheckDependencies=false
				}
				__PopCalUpdateSummaryValidator()
			}
			__PopCalSelectionChanged(o,_PopCal)
		}
	}
}

function __PopCalUpdateSummaryValidator()
{
	if(__PopCalTimerSummary==null)
	{
		__PopCalTimerSummary=window.setTimeout("__PopCalDisplaySummaryValidator()",250)
	}
}

function __PopCalDisplaySummaryValidator()
{
	clearTimeout(__PopCalTimerSummary)
	__PopCalTimerSummary=null
	if (typeof(ValidationSummaryOnSubmit)=="function")
	{
		if(typeof(Page_ValidationSummaries)!="undefined")
		{
			var i
			var _summary
			for(i=0;i<Page_ValidationSummaries.length;i++)
			{
				_summary=Page_ValidationSummaries[i]
				_summary.saveshowmessagebox=_summary.showmessagebox
				_summary.showmessagebox="False"
			}
			ValidatorUpdateIsValid()
			ValidationSummaryOnSubmit()
			for(i=0;i<Page_ValidationSummaries.length;i++)
			{
				_summary=Page_ValidationSummaries[i]
				_summary.showmessagebox=_summary.saveshowmessagebox
			}
		}
	}
	for(var i=0;i<__PopCalValidCalendarRanges.length;i++)
	{
		var o=document.getElementById(__PopCalValidCalendarRanges[i].Control)
		if(o)
		{
			var _v=document.getElementById(o.getAttribute("Calendar")+"_MessageError")
			if(_v)
			{
				if(_v.popupOverMessage)
				{
					if((_v.style.visibility=='hidden')||(_v.style.display=='none'))
					{
						if(_v.popupOverMessage.style.display!='none')
						{
							_v.popupOverMessage.style.display='none'
						}
					}
				}
			}
		}
	}
}

function __PopCalFormatControl(o)
{
	if(o.value!="")
	{
		var _PopCal=eval(o.getAttribute("Calendar"))
		if(_PopCal)
		{
			var _format=o.getAttribute("Format")
			var _Sep=__PopCalGetSeparator(o.value)
			var sRetVal=_PopCal.formatDate(o.value,__PopCalReplaceSeparators(_format,_Sep),_format)
			if(_format.indexOf("mmmm")!=-1)
			{
				if(sRetVal=="") sRetVal=_PopCal.formatDate(o.value,__PopCalReplaceSeparators(_format.replace("mmmm","mmm"),_Sep),_format)
				if(sRetVal=="") sRetVal=_PopCal.formatDate(o.value,__PopCalReplaceSeparators(_format.replace("mmmm","mm"),_Sep),_format)
			}
			else if(_format.indexOf("mmm")!=-1)
			{
				if(sRetVal=="") sRetVal=_PopCal.formatDate(o.value,__PopCalReplaceSeparators(_format.replace("mmm","mmmm"),_Sep),_format)
				if(sRetVal=="") sRetVal=_PopCal.formatDate(o.value,__PopCalReplaceSeparators(_format.replace("mmm","mm"),_Sep),_format)
			}
			else
			{
				if(sRetVal=="") sRetVal=_PopCal.formatDate(o.value,__PopCalReplaceSeparators(_format.replace("mm","mmmm"),_Sep),_format)
				if(sRetVal=="") sRetVal=_PopCal.formatDate(o.value,__PopCalReplaceSeparators(_format.replace("mm","mmm"),_Sep),_format)
			}
			if(sRetVal=="") sRetVal=_PopCal.formatDate(o.value,__PopCalReplaceSeparators("yyyy-mm-dd",_Sep),_format)
			o.value=sRetVal
			return(sRetVal!="")
		}
	}
	return(true)
}

function __PopCalShowMessageWaitForControl(_o)
{
	var o=document.getElementById(_o)
	if(o)
	{
		var _v=document.getElementById(o.getAttribute("Calendar")+"_MessageError")
		if (_v)
		{
			var _tl=PopCalGetTopLeft(o)
			if(_tl[0]==0)
			{
				window.setTimeout("__PopCalShowMessage('"+_o+"',null,null,true)",10)
			}
			else
			{
				__PopCalShowMessage(_o,null,null,true)
			}
			return
		}
	}
	window.setTimeout("__PopCalShowMessageWaitForControl('"+_o+"')",10)
}

function __PopCalShowMessage(_o,_m,_f,_w)
{
	var _focus=true
	var o=document.getElementById(_o)
	if(!o) return
	var _v=document.getElementById(o.getAttribute("Calendar")+"_MessageError")
	if(!_v) return
	if(_m) _v.innerHTML=_m
	if(_f!=null) _focus=_f
	if(!_w) __PopCalUpdateSummaryValidator()
	o.blankfield=true
	window.setTimeout("__PopCalBlankField('"+o.id+"')",500)
	if(_focus)
	{
		__PopCalControlFocus(o)
	}
	if(_v.getAttribute("ShowErrorMessage")=='false')
	{
		_v.style.visibility='hidden'
		_v.style.display='none'
		return
	}
	var _PopCal=eval(o.getAttribute("Calendar"))
	if(_PopCal)
	{
		if (_v.style.position.toLowerCase()=='absolute')
		{
			if((!_v.popupOverMessage)&&(_PopCal.ie)&&(_PopCal.ieVersion>=5.5))
			{
				if(document.body)
				{
					if(document.body.insertAdjacentHTML)
					{
						document.body.insertAdjacentHTML("afterBegin","<iframe id='popupOverMessage"+_PopCal.id+"' src='javascript:false;' scrolling=no frameborder=0 style='position:absolute;left:0px;top:0px;width:0px;height:0px;z-index:+10000;display:none;filter:progid:DXImageTransform.Microsoft.Alpha(opacity=0);'></iframe>")
						_v.popupOverMessage=document.getElementById("popupOverMessage"+_PopCal.id)
					}
				}
			}
		}
	}
	var _tl
	var c=document.getElementById(o.getAttribute("CalendarControl"))
	if(!c) c=document.getElementById(o.getAttribute("Calendar")+"_Control")
	if(!c) c=o
	if(_v.getAttribute("MessageAlignment")!='MessageContainer')
	{
		if(_v.getAttribute("MessageAlignment")=='RightCalendarControl')
		{
			_tl=PopCalGetTopLeft(c)
			_tl[0]-=(c.offsetHeight-1)
			_tl[1]+=(c.offsetWidth+10)
		}
		else
		{
			_tl=PopCalGetTopLeft(o)
			if((_v.style.padding=='2px')||(_v.style.padding=='2px 2px 2px 2px')) _tl[0]+=4
		}
		_v.style.top=(_tl[0]+1)+'px'
		_v.style.left=_tl[1]+'px'
	}
	_v.style.zIndex=+10000
	if(_v.getAttribute("ShowMessageBox")=='true')
	{
		_v.style.visibility='hidden'
		_v.style.display='none'
		if(_v.innerText)
		{
			alert(_v.innerText)
		}
		else
		{
			alert(_v.innerHTML)
		}
	}
	else
	{
		if(_v.getAttribute("MessageAlignment")=='MessageContainer')
		{
			_v.style.visibility='visible'
			if (_PopCal.lr==0)
			{
				if (_v.style.position.toLowerCase()=='absolute')
				{
					if((_v.getAttribute("rtlLeft")!="")&&(_v.getAttribute("rtlLeft")!=null))
					{
						var _Right=-(_v.offsetWidth)
						_Right+=parseInt(_v.getAttribute("rtlLeft"),10)
						_Right+=parseInt(_v.getAttribute("rtlWidth"),10)
						_v.style.left=_Right+'px'
					}
				}
			}
		}
		else
		{
			_v.style.display=''
			if (_PopCal.lr==0)
			{
				if(_v.getAttribute("MessageAlignment")=='RightCalendarControl')
				{
					_tl=PopCalGetTopLeft(c)
					_v.style.left=(_tl[1]-(c.offsetWidth+_v.offsetWidth)+10)+'px'
				}
				else
				{
					_tl=PopCalGetTopLeft(o)
					_v.style.left=(_tl[1]+o.offsetWidth-_v.offsetWidth)+'px'
				}
			}
		}
		if(_v.popupOverMessage)
		{
			_v.popupOverMessage.style.top=parseInt(_v.style.top,10)+'px'
			_v.popupOverMessage.style.left=parseInt(_v.style.left,10)+'px'
			_v.popupOverMessage.style.height=_v.offsetHeight+'px'
			_v.popupOverMessage.style.width=_v.offsetWidth+'px'
			_v.popupOverMessage.style.display=''
		}
	}
}

function __PopCalBlankField(_o)
{
	var o=document.getElementById(_o)
	if(!o) return
	o.value=""
	o.blankfield=null
	if((o.valueOnFocus!=null)&&(o.valueOnFocus!=''))
	{
		var _PopCal=eval(o.getAttribute("Calendar"))
		__PopCalSelectionChanged(o,_PopCal)
	}
	o.valueOnFocus=null
}

function __PopCalSetErrorMessage(_ValidControl,_DateMsg,_textMessage)
{
	if(_ValidControl)
	{
		_ValidControl.errormessage=_DateMsg
		if(_textMessage!='')
		{
			return(_textMessage)
		}
	}
	return(_DateMsg)
}

function __PopCalGetYYYYMMDD(o)
{
	if(o)
	{
		var _PopCal=eval(o.getAttribute("Calendar"))
		if(_PopCal)
		{
			return(_PopCal.formatDate(o.value,o.getAttribute("Format"),"yyyy-mm-dd"))
		}
	}
	return("")
}

function __PopCalGetFromYYYYMMDD(o)
{
	var _DateFrom=""
	if(o)
	{
		if(__PopCalTemporal.indexOf(","+o.id.toLowerCase()+",")!=-1) return(_DateFrom)
		__PopCalTemporal+=(o.id.toLowerCase()+",")
		var _PopCal=eval(o.getAttribute("Calendar"))
		for(var i=0;i<__PopCalValidCalendarRanges.length;i++)
		{
			var _Range=__PopCalValidCalendarRanges[i]
			if(_Range.Control==o.id)
			{
				if(_Range.FromRange=="Hoy")
				{
					_DateFrom=_PopCal.formatDate("Hoy","yyyy-mm-dd","yyyy-mm-dd")
				}
				else if(_Range.FromRange.substr(0,2)=="C:")
				{
					var _From=document.getElementById(_Range.FromRange.substr(2))
					if (!_From.blankfield)
					{
						_DateFrom=__PopCalGetYYYYMMDD(_From)
						if(_DateFrom=="")
						{
							_DateFrom=__PopCalGetFromYYYYMMDD(_From)
						}
					}
				}
				else
				{
					_DateFrom=_Range.FromRange
				}
				if(_DateFrom!="")
				{
					_DateFrom=_PopCal.addDays(_DateFrom,"yyyy-mm-dd",_Range.FromIncrement)
				}
				break
			}
		}
	}
	return(_DateFrom)
}

function __PopCalGetToYYYYMMDD(o)
{
	var _DateTo=""
	if(o)
	{
		if(__PopCalTemporal.indexOf(","+o.id.toLowerCase()+",")!=-1) return(_DateTo)
		__PopCalTemporal+=(o.id.toLowerCase()+",")
		var _PopCal=eval(o.getAttribute("Calendar"))
		for(var i=0;i<__PopCalValidCalendarRanges.length;i++)
		{
			var _Range=__PopCalValidCalendarRanges[i]
			if(_Range.Control==o.id)
			{
				if(_Range.ToRange=="Hoy")
				{
					_DateTo=_PopCal.formatDate("Hoy","yyyy-mm-dd","yyyy-mm-dd")
				}
				else if(_Range.ToRange.substr(0,2)=="C:")
				{
					var _To=document.getElementById(_Range.ToRange.substr(2))
					if (!_To.blankfield)
					{
						_DateTo=__PopCalGetYYYYMMDD(_To)
						if(_DateTo=="")
						{
							_DateTo=__PopCalGetToYYYYMMDD(_To)
						}
					}
				}
				else
				{
					_DateTo=_Range.ToRange
				}
				if(_DateTo!="")
				{
					_DateTo=_PopCal.addDays(_DateTo,"yyyy-mm-dd",_Range.ToIncrement)
				}
				break
			}
		}
	}
	return(_DateTo)
}

function __PopCalObjectCalendarRange()
{
	this.Control=""
	this.FromRange=""
	this.FromIncrement=0
	this.FromMessage=""
	this.ToRange=""
	this.ToIncrement=0
	this.ToMessage=""
}

function __PopCalAddCalendarRange(_Control,_FromRange,_FromIncrement,_FromMessage,_ToRange,_ToIncrement,_ToMessage)
{
	var _Range=new __PopCalObjectCalendarRange()
	_Range.Control=_Control
	_Range.FromRange=_FromRange
	_Range.FromIncrement=_FromIncrement
	_Range.FromMessage=_FromMessage
	_Range.ToRange=_ToRange
	_Range.ToIncrement=_ToIncrement
	_Range.ToMessage=_ToMessage
	if(_Range.FromRange=='') _Range.FromRange='1900-01-01'
	if(_Range.ToRange=='') _Range.ToRange='2099-12-31'
	var _idx=__PopCalValidCalendarRanges.length
	for(var i=0;i<__PopCalValidCalendarRanges.length;i++)
	{
		if(__PopCalValidCalendarRanges[i].Control==_Range.Control)
		{
			_idx=i
			break
		}
	}
	__PopCalValidCalendarRanges[_idx]=_Range
}

function __PopCalGetSeparator(_f)
{
	if(_f.indexOf("/")!=-1) return("/")
	else if(_f.indexOf("-")!=-1) return("-")
	return(".")
}

function __PopCalReplaceSeparators(_f,_s)
{
	var _r=_f
	_r=_r.split('-').join(_s)
	_r=_r.split('/').join(_s)
	_r=_r.split('.').join(_s)
	return(_r)
}

function __PopCalControlFocus(o)
{
	if(o.getAttribute("SetFocusOnError")=="false") return
	if(__PopCalLastControlFocus!="") return
	if(o.id)
	{
		__PopCalLastControlFocus=o.id
		window.setTimeout("__PopCalWaitForSetControlFocus()",250)
	}
	else
	{
		window.setTimeout("__PopCalLastControlFocus=''",500)
		try
		{
			o.focus()
		}
		catch (e)
		{
		}
	}
}

function __PopCalWaitForSetControlFocus()
{
	if (__PopCalLastControlFocus!="")
	{
		var o=document.getElementById(__PopCalLastControlFocus)
		window.setTimeout("__PopCalLastControlFocus=''",500)
		if(!o) return
		try
		{
			o.focus()
		}
		catch (e)
		{
		}
	}
}

function __PopCalValueTrim(s)
{
	var m=s.match(/^\s*(\S+(\s+\S+)*)\s*$/)
	return ((m==null)?"":m[1])
}

function __PopCalCustomValidatorEvaluateIsValid(val)
{
	var value=""
	if (typeof(val.controltovalidate)=="string")
	{
		var o=document.getElementById(val.controltovalidate)
		if(o)
		{
			if(o.valueOnFocus==null) o.valueOnFocus=o.value
			value=o.value
		}
	}
	var args={Value:value,IsValid:true}
	if(typeof(val.clientvalidationfunction)=="string")
	{
		eval(val.clientvalidationfunction+"(val,args);")
	}
	return args.IsValid
}

function __PopCalSelectionChanged(_TextBox,_PopCal)
{
	var _retval=true
	if(_PopCal.ClientScriptOnDateChanged!='')
	{
		if(typeof(eval("window."+_PopCal.ClientScriptOnDateChanged))=="function")
		{
			_retval=eval("window."+_PopCal.ClientScriptOnDateChanged+"(_TextBox,_PopCal)")
			_retval=(typeof(_retval)=='boolean')?_retval:true
		}
	}
	if(_retval)
	{
		if(_TextBox.getAttribute("Buffer")=="true")
		{
			eval(_TextBox.getAttribute("PostBack").toString().replace('9999x99x99',_TextBox.value))
		}
	}
}

function __PopCalGetTicks()
{
	return (new Date())-(new Date(2005,0,1))
}