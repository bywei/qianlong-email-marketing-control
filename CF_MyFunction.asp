
<%'以下为公用函数
Function goback(ByVal str,ByVal alertstr) '为空时后退
 if str="" then 
  response.write "<script>" 
  response.write "alert('"&alertstr&"');" 
  response.write "history.go(-1)" 
  response.write "</script>" 
  Call ConnClose()
  response.end 
 else
  goback=str
 end if
End Function

Function alertback(ByVal alertstr,ByVal backnum) 
  response.write "<script>" 
  response.write "alert('"&alertstr&"');" 
  response.write "history.go(-"&backnum&")" 
  response.write "</script>" 
  Call ConnClose()
  response.end 
End Function

Function AlertUrl(ByVal alertstr,ByVal url) 
  response.write "<script>" 
  response.write "alert('"&alertstr&"');" 
  response.write "location.href='"&url&"';" 
  response.write "</script>" 
  Call ConnClose()
  response.end 
End Function

Function AlertClose(ByVal alertstr)
  response.write "<script>"
  response.write "alert('"&alertstr&"');"
  response.write "window.close();"
  response.write "</script>"
  Call ConnClose()
  response.end
End Function

Function GotoUrl(ByVal url) 
  response.write "<script>" 
  response.write "location.href='"&url&"';" 
  response.write "</script>" 
  Call ConnClose()
  response.end 
End Function

    Function CheckInput_Letter(ByVal InputStr) '检查用户名输入的合法性
        CheckInput_Letter = -1

        For I = 1 To Len(InputStr)
            C = LCase(Mid(InputStr, I, 1)) '------------分割成每个字母或数字------------------
            If InStr("abcdefghijklmnopqrstuvwxyz_", C) <= 0 And Not IsNumeric(C) Then
                CheckInput_Letter = 0
                Exit For
            End If
        Next
    End Function

function checkinput_blank(ByVal inputstr) '检查密码输入的合法性
for i = 1 to Len(inputstr)
 c = Lcase(Mid(inputstr, i, 1)) '------------分割成每个字母或数字------------------
  if InStr(" ", c) > 0 or InStr("　", c) > 0 then
  response.write "<script language='javascript'>" & VbCRlf
  response.write "alert('请不要输入空格!');" & VbCrlf
  response.write "history.go(-1);" & vbCrlf
  response.write "</script>" & VbCRLF
  Call ConnClose()
  response.end
  end if
next
 checkinput_blank=inputstr
End Function

Function ChkStr(ByVal ParaValue,ByVal ParaType)'参数类型-数字型(1是字符，2是数字，3是日期
 If ParaType = 1 then
  ChkStr = Replace(ParaValue,"'","''")
 ElseIf ParaType = 2 then
  If ParaValue="" Then
   ChkStr = ParaValue
  Else
   If IsNumeric(ParaValue) = True Then
    If Abs(ParaValue)<=214748367 Then
	 ChkStr = ParaValue
	Else
	 Response.Write "以下传递的参数类型有错误！<br>" & ParaValue
     Response.End
	End If
   Else   
    Response.Write "以下传递的参数类型有错误！<br>" & ParaValue
    Response.End
   End If
  End If
 ElseIf ParaType = 3 then
  If ParaValue<>"" And (IsDate(ParaValue) = False) then
   Response.Write "以下传递的参数类型有错误！<br>" & ParaValue
   Response.End
  Else
   ChkStr = ParaValue
  End If
 End If
End Function

Function GetFieldValues(ByVal TalbeName,ByVal FieldNmae,ByVal Fieldvalues,ByVal FieldType,ByVal GetFieldName)  '通用,通过一个表的字段，得到表中某个字段的值
 if FieldType=1 then
	sql="select "&GetFieldName&" from "& TalbeName &" where "& FieldNmae &"="& Fieldvalues
 elseif FieldType=2 then
	sql="select "&GetFieldName&" from "& TalbeName &" where "& FieldNmae &"='"& Fieldvalues &"'"
 end if
 set FieldValues=server.createobject("adodb.recordset")
 FieldValues.open sql,conn,1,1
	if Not FieldValues.eof then
	 GetFieldValues=FieldValues(0)
	end if
	FieldValues.close
End function

Function MyRate(ByVal snum,ByVal bnum)

If snum>0 Then
 MyRate=formatnumber(snum/bnum*100,2,-1)
Else
 MyRate=0
End If

End Function

Function HttpPath(ByVal Assort)
Ser=Request.servervariables("SERVER_NAME")
Scr=Request.servervariables("SCRIPT_NAME")
Port=Request.Servervariables("SERVER_PORT")

Scr_2=StrReverse(Mid(StrReverse(Scr),Instr(StrReverse(Scr),"/")))
If Assort=1 Then
 HttpPath=Ser
 
ElseIf Assort=2 Then
 If Port="80" Then
  HttpPath="http://"&Ser&Scr_2
 Else
  HttpPath="http://"&Ser&":"&Port&Scr_2
 End If
 
ElseIf Assort=3 Then
 If Port="80" Then
  HttpPath="http://"&Ser&Scr
 Else
  HttpPath="http://"&Ser&":"&Port&Scr
 End If
End If
End Function



Function PxFilter(ByVal px,ByVal pxok)
px=Lcase(px)
pxok=Lcase(pxok)

PxArrary=Split(Pxok,",")

For I= 0 To Ubound(PxArrary)
  If PxArrary(I)=Px Then J=1
Next

If J<>1 Then Call Alertback("禁止此类排序",1)

End Function


Function BreakUrl(ByVal Url,ByVal BreakType)
Url=Lcase(Url)
If Url<>"" Then
 UrlArrary=Split(Url,"/")
 UrlHead=UrlArrary(2)
 UrlTail=UrlArrary(Ubound(UrlArrary))


 If BreakType=1 Then
  BreakUrl=UrlHead
 ElseIf BreakType=2 Then
  If UrlTail<>"" Then
   BreakUrl=UrlTail
  Else
   BreakUrl=UrlHead
  End if
 End if
Else
 BreakUrl=Url
End if

End Function

Function GetSearchKeyword(ByVal Url, ByVal KeyWordFlag)
 If InStr(Url, "?" & KeyWordFlag & "=") > 0 Then
  KeyWordFlag = "?" & KeyWordFlag & "="
 ElseIf InStr(Url, "&" & KeyWordFlag & "=") > 0 Then
  KeyWordFlag = "&" & KeyWordFlag & "="
 End If

 UrlArrary = Split(Url, KeyWordFlag)

 If UBound(UrlArrary) > 0 Then
  UrlTail = UrlArrary(1)
  If InStr(UrlTail, "&") = 0 Then
   GetSearchKeyword = Mid(UrlTail, 1, 1000)
  Else
   GetSearchKeyword = Mid(UrlTail, 1, InStr(UrlTail, "&") - 1)
  End If
 End If 
End Function

Function GetSysSet(ByVal SiteNum)
 MyArrary=Split(Application("CFCountSysSet"),"*")
 GetSysSet=MyArrary(SiteNum)
End Function


Function GetTurnTime(byval Num)
Num=Cstr(Num)
If Len(Num)=1 Then
 GetTurnTime="0"&Num
Else
 GetTurnTime=Num
End if
End Function


Function URLDecode(byval enStr)
  dim deStr
  dim c,i,v
  deStr=""
  for i=1 to len(enStr)
  c=Mid(enStr,i,1)
  if c="%" then
  v=eval("&h"+Mid(enStr,i+1,2))
  if v<128 then
  deStr=deStr&chr(v)
  i=i+2
  else
  if isvalidhex(mid(enstr,i,3)) then
  if isvalidhex(mid(enstr,i+3,3)) then
  v=eval("&h"+Mid(enStr,i+1,2)+Mid(enStr,i+4,2))
  deStr=deStr&chr(v)
  i=i+5
  else
  v=eval("&h"+Mid(enStr,i+1,2)+cstr(hex(asc(Mid(enStr,i+3,1)))))
  deStr=deStr&chr(v)
  i=i+3 
  end if 
  else 
  destr=destr&c
  end if
  end if
  else
  if c="+" then
  deStr=deStr&" "
  else
  deStr=deStr&c
  end if
  end if
  next
  URLDecode=deStr
  end function

  function isvalidhex(str)
  isvalidhex=true
  str=ucase(str)
  if len(str)<>3 then isvalidhex=false:exit function
  if left(str,1)<>"%" then isvalidhex=false:exit function
  c=mid(str,2,1)
  if not (((c>="0") and (c<="9")) or ((c>="A") and (c<="Z"))) then isvalidhex=false:exit function
  c=mid(str,3,1)
  if not (((c>="0") and (c<="9")) or ((c>="A") and (c<="Z"))) then isvalidhex=false:exit function
  end function
  
function GetUnicode(Str) 
  dim i 
  dim Str_one 
  dim Str_unicode 
  if(isnull(Str)) then
     exit function
  end if
  for i=1 to len(Str) 
    Str_one=Mid(Str,i,1) 
    Str_unicode=Str_unicode&chr(38) 
    Str_unicode=Str_unicode&chr(35) 
    Str_unicode=Str_unicode&chr(120) 
    Str_unicode=Str_unicode& Hex(ascw(Str_one)) 
    Str_unicode=Str_unicode&chr(59) 
  next 
  GetUnicode=Str_unicode 
end function 

'解码出始
function UTF2GB(byval UTFStr) 
for Dig=1 to len(UTFStr) 
if mid(UTFStr,Dig,1)="%" then 
if len(UTFStr) >= Dig+8 then 
GBStr=GBStr & ConvChinese(mid(UTFStr,Dig,9)) 
Dig=Dig+8 
else 
GBStr=GBStr & mid(UTFStr,Dig,1) 
end if 
else 
GBStr=GBStr & mid(UTFStr,Dig,1) 
end if 
next 
UTF2GB=GBStr 
end function 

function ConvChinese(x) 
A=split(mid(x,2),"%") 
i=0 
j=0 

for i=0 to ubound(A) 
A(i)=c16to2(A(i)) 
next 

for i=0 to ubound(A)-1 
DigS=instr(A(i),"0") 
Unicode="" 
for j=1 to DigS-1 
if j=1 then 
A(i)=right(A(i),len(A(i))-DigS) 
Unicode=Unicode & A(i) 
else 
i=i+1 
A(i)=right(A(i),len(A(i))-2) 
Unicode=Unicode & A(i) 
end if 
next 

if len(c2to16(Unicode))=4 then 
ConvChinese=ConvChinese & chrw(int("&H" & c2to16(Unicode))) 
else 
ConvChinese=ConvChinese & chr(int("&H" & c2to16(Unicode))) 
end if 
next 
end function 

function c2to16(x) 
i=1 
for i=1 to len(x) step 4 
c2to16=c2to16 & hex(c2to10(mid(x,i,4))) 
next 
end function 

function c2to10(x) 
c2to10=0 
if x="0" then exit function 
i=0 
for i= 0 to len(x) -1 
if mid(x,len(x)-i,1)="1" then c2to10=c2to10+2^(i) 
next 
end function 

function c16to2(x) 
i=0 
for i=1 to len(trim(x)) 
tempstr= c10to2(cint(int("&h" & mid(x,i,1)))) 
do while len(tempstr)<4 
tempstr="0" & tempstr 
loop 
c16to2=c16to2 & tempstr 
next 
end function 

function c10to2(x) 
mysign=sgn(x) 
x=abs(x) 
DigS=1 
do 
if x<2^DigS then 
exit do 
else 
DigS=DigS+1 
end if 
loop 
tempnum=x 

i=0 
for i=DigS to 1 step-1 
if tempnum>=2^(i-1) then 
tempnum=tempnum-2^(i-1) 
c10to2=c10to2 & "1" 
else 
c10to2=c10to2 & "0" 
end if 
next 
if mysign=-1 then c10to2="-" & c10to2 
end function 
'解码结束

Function GetIpArea(ByVal Ip)
inIP=Ip
inIPs=split(inIP,".")
inIPnum=inips(0)*256*256*256 + inips(1)*256*256 + inips(2)*256 + inips(3)


Sql="Select Area,Area_2,Area_3 from CFCount_IpAddress where Ip_1<="&inIPnum&" and Ip_2>="&inIPnum
set RsIp=Conn.Execute(Sql)

If Not RsIp.Eof Then
 GetIpArea= RsIp("Area")&"-"&RsIp("Area_2")&"-"&RsIp("Area_3")
Else
  GetIpArea= "<a href='http://js.tb11.com' target=_blank>未知地区，点击下载最新IP数据库</a>"
End If
RsIp.Close
End Function

Function GetPicStyle(ByVal UserName)
Set RsPicStyle= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
RsPicStyle.Open Sql,Conn,1,1
If RsPicStyle("ShowType")=1 Then
 Counter=RsPicStyle("ShowTotal")
ElseIf RsPicStyle("ShowType")=2 Then
 Counter=RsPicStyle("RealShowTotal")
ElseIf RsPicStyle("ShowType")=3 Then
 Counter=RsPicStyle("RealIpTotal")
End if   
CountLen=Len(Counter)
ZeroNum=RsPicStyle("PicNum")-CountLen

For I=1 To ZeroNum
 LinkUrl="CounterPic/"&RsPicStyle("Style")&"/0"
 CounterHtm=CounterHtm&"<img src="&LinkUrl&".gif border='0'>"
Next
For I=1 To CountLen
 Pic=Mid(Counter,I,1)
 LinkUrl="CounterPic/"&RsPicStyle("Style")&"/"&Pic
 CounterHtm=CounterHtm&"<img src="&LinkUrl&".gif border='0'>"
Next
GetPicStyle=CounterHtm
RsPicStyle.Close
End Function

Function GetAppChange(ByVal UserName,ByVal IP,ByVal CFCountVisitTotal,ByVal Currweb)
 MyArray=Split(Application("CFCountLy_"&UserName),"|")
 For I=0 To Ubound(MyArray)-1
  If Instr(MyArray(I),IP)=0 Then
   AllStr=AllStr&MyArray(I)&"|"
  Else
   MyArray_2=Split(MyArray(I),"\")
   MyStr=Ip&"\"&CFCountVisitTotal&"\"&MyArray_2(2)&"\"&MyArray_2(3)&"\"&Now()&"\"&CurrWeb&"\"&Int(MyArray_2(6))+1&"|"
   AllStr=AllStr&MyStr
  End if
 Next
   
 GetAppChange=AllStr
End Function

Function GetAppChange_2(ByVal UserName,ByVal IP,ByVal CFCountVisitTotal,ByVal Currweb,ByVal OnlineTime)
 MyArray=Split(Application("CFCountOnline_"&UserName),"|")
 
 If Ubound(Myarray)>1000 Then K=1'只保留1000条记录
 J=0
 For I=K To Ubound(MyArray)-1
  MyArray_2=Split(MyArray(I),"\")
  If Instr(MyArray(I),IP)>0 And J=0 Then   
   MyStr=Ip&"\"&CFCountVisitTotal&"\"&MyArray_2(2)&"\"&MyArray_2(3)&"\"&Now()&"\"&CurrWeb&"\"&Int(MyArray_2(6))+1&"|"
   AllStr=AllStr&MyStr
   J=1
  Else
   If Ubound(MyArray_2)=6 Then
    If DateDiff("n",Cdate(MyArray_2(4)),Now())<=OnlineTime Then AllStr=AllStr&MyArray(I)&"|"
   End If
  End if
 Next
   
 GetAppChange_2=AllStr
End Function

Function GetWeek(ByVal Mydate)
weeknum = weekDay(Mydate)
select case weeknum
 case "1"
  data="星期天"
 case "2"
  data="星期一"
 case "3"
  data="星期二"
 case "4"
  data="星期三"
 case "5"
  data="星期四"
 case "6"
  data="星期五"
 case "7"
  data="星期六"
 case else
  data="错误的时间"
end select
GetWeek=data
End Function

Function GetMySet(ByVal MyStr)

 If IsEmpty(Application("CFCount_MySet")) Then
  Sql = "Select OtherSet From CFCount_Admin"
  Set Rs_MySet = Conn.Execute(Sql)
  Application("CFCount_MySet") = Rs_MySet("OtherSet")
  Rs_MySet.Close
 End If
 
 MyArray_MySet = Split(Application("CFCount_MySet"), "|||")
 For I_MySet = 0 To UBound(MyArray_MySet)
  If LCase(Left(MyArray_MySet(I_MySet), Len(MyStr) + 1)) = LCase(MyStr) & "=" Then GetMySet = Mid(MyArray_MySet(I_MySet), Len(MyStr) + 2)
 Next

End Function


Function GetSkinColor(byval SkinColor,byval Assort)
 MyArray_SkinColor=Split(SkinColor,"|")
 If Assort<=Ubound(MyArray_SkinColor) Then
  GetSkinColor=MyArray_SkinColor(Assort)
 End If
End Function

Function ConnClose()
 If IsObject(Rs)=True Then
  Rs.Close
  Set Rs=Nothing
 End If

 If IsObject(Rs2)=True Then
  Rs2.Close
  Set Rs2=Nothing
 End If

 If IsObject(RsUser)=True Then
  RsUser.Close
  Set RsUser=Nothing
 End If

 If IsObject(RsSet)=True Then
  RsSet.Close
  Set RsSet=Nothing
 End If

 If IsObject(Conn)=True Then
  Conn.Close
  Set Conn=Nothing
 End If
End Function
%>