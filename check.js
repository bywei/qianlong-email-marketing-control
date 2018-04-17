//----------------isNumeric(str)函数------------------------
function isNumeric(str)
{
	var i=0;
	var j=0;
	while(i<=str.length-1){
		str1=str.charAt(i);
		if (str1=='.'){
			if ((j>0)||(i==0)||(i==(str.length-1))||(i>=2&&i<(str.length-1)&&parseInt(str.substring(0,i))==0))
				return false;
			j=j+1;
		}
		else if(isNaN(parseInt(str1))==true) return false;
		i=i+1;
	}
	return true;
}	

function window.confirm(str)
{
    execScript("n = msgbox('"+ str +"', 257)", "vbscript");
    return(n == 1);
}
//strDate = "2004-1-2"
//return  "2004-01-02"
// 在每个"-"字符处进行分解。
function FormatDate0(strDate)
{
var strary, yy,mm,dd;
strary = strDate.split("-");
yy = strary[0];
mm = (strary[1].length==1)?"0"+strary[1]:strary[1];
dd = (strary[2].length==1)?"0"+strary[2]:strary[2];
strsplit = "-"
   return(yy + strsplit + mm + strsplit + dd );
}

//---------判断文本框内容是否为空 +删除空格的字符 -------------
function checkNull(sTextBox,TextWord)
{
	sTextBoxValue1=eval("document." + sTextBox + ".value");
	sTextBoxValue=Javatrim1(sTextBoxValue1) //删除空格的字符
	if(sTextBoxValue.length==0)
		{
			alert( TextWord+"不能为空！");
			eval("document." + sTextBox + ".focus()");
			return 1;
		}
}

function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
//判断是输入栏否为空
//object_name 为对象名称，tishi=1 为是否显示对话框,word 为提示语句,kongge=1 为去除空格,
function IsNull(object_name,tishi,word,kongge)
{
	var string;
	string=new String(object_name);
	if (kongge==1)
	{string=Javatrim1(string);} //删除空格的字符 
	//alert("返回的字符集="+string+"长度为="+string.length)
	if (string.length==0)
	{
		if (tishi==1)
		{
		alert(word);
		}
		return false;
	}
	return true;
}
   
//删除字符开头和结尾的空格
function Javatrim1(str){
	var i=0;
	var j;
	var len=str.length;

	trim1str="";
	if(j<0) return trim1str;
	flagbegin= true;
	flagend= true;
		
	while (flagbegin== true){
		if (str.charAt(i)==" "){
			i++;
			flagbegin=true;
		}
		else
		{
			
			flagbegin=false;
		}
	} 
	//前面有i个空格
	j=len-1;
	var k=0;
	while (flagend==true)
	{
		if (str.charAt(j)==" ")
		{
			j--;
			flagend=true;
			k++;
		}
		else{
			flagend=false;
		}
	}
        
	//后面有k个空格
	//alert('前面有'+i+'个空格！');
	//alert('后面有'+k+'个空格！');
	if (str.length==i)
	{
	 //alert("您的输入全为空格！")
	 trim1str="";
	 return trim1str;
	}
	trim1str=str.substring(i,j+1);
	//alert("bf"+trim1str+"fb");
	return trim1str;
}

// 判断电子邮件是否格式正确
function IsEmail(object_name)
{
	var string;
	string=new String(object_name);
	var len=string.length;
	if (string.indexOf("@",1)==-1||string.indexOf(".",1)==-1||string.length<7)
		{
		alert("电子邮箱的格式不对，请重新填写！");
		return false;
		}
	if (string.charAt(len-1)=="."||string.charAt(len-1)=="@")
		{
		alert("电子邮箱的格式不对，请重新填写！");
		return false;
		}
	return true;
}
   

//判断输入栏的最小和最大长度是否越界
function OverLength(object_name,max,min,max_word,min_word,kongge) //kongge=1 为处理掉字符串中的空格
{
	var string;
	string=new String(object_name);
	if (kongge==1)
		{
		 string=Javatrim1(string);
		}
	if (string.length>max)
		{
		 alert(max_word);
		 return false;
		}
	if (string.length<min)
		{
		 alert(min_word);
		 return false;
		}
}
function IsWhitespace (s) //是否包涵空格
{  
  	var whitespace = " \t\n\r";
 	var i;
  	for (i = 0; i < s.length; i++)
  	{   
   	   var c = s.charAt(i);
   	   if (whitespace.indexOf(c) >= 0) 
	   {
		  return true;
		}
  	}

    return false;
}

function IsPassword(strPassword)
{
	if(strPassword=="") return false;
	var lngLength = strPassword.length;
	var strCharSet="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
	for(i=0; i<lngLength; i++)
	{
		if (strCharSet.search(strPassword.substr(i,1))<0) return false;
	}
	return true;
}

function IsPostcode(object_name)
{
    var string;
    string=new String(object_name);
    if (string.length!=6)
    {
     alert("邮政编码应为6位数字，您的输入有误，请重新填写！");
     return false;
     }
     
    if (isNaN(string))
    {
     alert("邮政编码应为数字，您的输入有误，请重新填写！");
     return false;
     }
    return true;
}



