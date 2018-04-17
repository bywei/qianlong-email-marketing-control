//全选按钮
function CheckAllCK1()
{
	for(i=0;i<document.myform.elements.length;i++)
	{
		if(document.myform.elements[i].checked){
			document.myform.elements[i].checked=false;
			document.myform.CheckAll.checked=false;
		}
		else{
			document.myform.elements[i].checked = true;
			document.myform.CheckAll.checked = true;
		}
	}
}

function CheckAllCK()
{
	for(i=0;i<document.myform.ck.length;i++)
	{
		if(document.myform.CheckAll.checked){
			document.myform.ck[i].checked=true;
			document.myform.CheckAll.checked=true;
		}
		else{
			document.myform.ck[i].checked = false;
			document.myform.CheckAll.checked = false;
			
		}
	}
}
function LoadFinish(){ 
　　document.all('LoadProcess').style.visibility = "hidden"; 
} 

 //日历效果
function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
	if (ReturnStr!='007007007007') SetObj.value=ReturnStr;
	return ReturnStr;
}
