//Open Modal Window
function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:yes;help:no;scroll:yes;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
}
//awen add   selectAll(this.form) ���ü���
function selectAll(f)
{
	for(i=0;i<f.length;i++)
	{
		if(f(i).type=="checkbox" && f(i)!=event.srcElement)
		{
			f(i).checked=event.srcElement.checked;
		}
	}
}
function CheckNumber(Obj,DescriptionStr)
{
	if (Obj.value!='' && (isNaN(Obj.value) || Obj.value<0))
	{
		alert(DescriptionStr+"Ӧ����Ч���֣�");
		Obj.value="";
		Obj.focus();
	}
}
function OpenWindow(Url,Width,Height,WindowObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:yes;');
	return ReturnStr;
}
//Open Modal Window
function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:yes;');
	//	var ReturnStr=OpenEditerWindow(Url,WindowObj,'scrollbars=0,resizable=1,top=50,left=50,width='+Width+',height='+Height);
	if (ReturnStr!='007007007007') SetObj.value=ReturnStr;
	return ReturnStr;
}
//Open Editer Window
function OpenEditerWindow(Url,WindowName,Width,Height)
{
	window.open(Url,WindowName,'toolbar=0,location=0,maximize=1,directories=0,status=1,menubar=0,scrollbars=1,resizable=1,top=50,left=50,width='+Width+',height='+Height);
}
//CSS��������
function overColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="hback_1"
		Obj.bgColor="";//��ɫҪ��
	}
	
}
function outColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="hback";
		Obj.bgColor="";
	}
}

//===========������Ϣ��ʾ���������������� coder by Terry Wen
function ShowErr(msg){
	alert(msg);
	var type = typeof(window.dialogArguments);
	var openerType = typeof(window.opener);
	if ( type != 'undefined' && openerType == 'undefined' )
	{
		//window.dialogArguments.location.reload();
		CloseWin();
	}
	else
	{
		if (document.referrer=="")
			CloseWin();
		else
			window.location.href=document.referrer;
	}
}
//===========�رմ��ڵĺ���(��������) coder by Terry Wen
function CloseWin()
{
	var ua=navigator.userAgent
	var ie=navigator.appName=="Microsoft Internet Explorer"?true:false
	if(ie)
	{
		var IEversion=parseFloat(ua.substring(ua.indexOf("MSIE ")+5,ua.indexOf(";",ua.indexOf("MSIE "))))
		if(IEversion< 5.5)
		{
			var str  = '<object id=noTipClose classid="clsid:ADB880A6-D8FF-11CF-9377-00AA003B7A11">'
			str += '<param name="Command" value="Close"></object>';
			document.body.insertAdjacentHTML("beforeEnd", str);
			document.all.noTipClose.Click();
		}
		else
		{
			window.opener =null;
			window.close();
		}
	}
	else
	{
		window.close()
	}
}
