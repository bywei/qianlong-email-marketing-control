<!-- #include file="mysetting.asp" -->
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 Walle 其它版本的Z-blog未知
'// 插件制作:    kikin(http://www.2sos.net/)
'// 备    注:    在文章前增加固定内容，可用于挂广告和版权说明 - 挂口页
'// 最后修改：   2011-10-23
'// 最后版本:    0.5
'///////////////////////////////////////////////////////////////////////////////
'常量标记

'注册插件
Call RegisterPlugin("AddTextF","ActivePlugin_AddTextF")
Function ActivePlugin_AddTextF() 
'挂上接口
  Call Add_Filter_Plugin("Filter_Plugin_TArticle_Export_TemplateTags","AddTextF_Main")
End Function
Function AddTextF_Main(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)
  If aryTemplateTagsValue(2)>2 Then

    If sTagInf <>"6" then  aryTemplateTagsValue(5) = AddTextF_Key(aryTemplateTagsValue(5)) 
     
    If sPosA= "Before"  Then  aryTemplateTagsValue(5) =AddTextF_Before(sMyWordA,aryTemplateTagsValue(5),sHPosA)
    If sPosA="After"  Then  aryTemplateTagsValue(5) =AddTextF_After(sMyWordA,aryTemplateTagsValue(5),sHPosA)
    If sPosA="Head"  Then  aryTemplateTagsValue(5) =AddTextF_Head(sMyWordA,aryTemplateTagsValue(5),sHPosA)
    If sPosA="Foot"  Then  aryTemplateTagsValue(5) =AddTextF_Foot(sMyWordA,aryTemplateTagsValue(5),sHPosA)
    
    If sPosB= "Before"  Then  aryTemplateTagsValue(5) =AddTextF_Before(sMyWordB,aryTemplateTagsValue(5),sHPosB)
    If sPosB="After"  Then  aryTemplateTagsValue(5) =AddTextF_After(sMyWordB,aryTemplateTagsValue(5),sHPosB)
    If sPosB="Head"  Then  aryTemplateTagsValue(5) =AddTextF_Head(sMyWordB,aryTemplateTagsValue(5),sHPosB)
    If sPosB="Foot"  Then  aryTemplateTagsValue(5) =AddTextF_Foot(sMyWordB,aryTemplateTagsValue(5),sHPosB)

    If sPosC= "Before"  Then  aryTemplateTagsValue(5) =AddTextF_Before(sMyWordC,aryTemplateTagsValue(5),sHPosC)
    If sPosC="After"  Then  aryTemplateTagsValue(5) =AddTextF_After(sMyWordC,aryTemplateTagsValue(5),sHPosC)
    If sPosC="Head"  Then  aryTemplateTagsValue(5) =AddTextF_Head(sMyWordC,aryTemplateTagsValue(5),sHPosC)
    If sPosC="Foot"  Then  aryTemplateTagsValue(5) =AddTextF_Foot(sMyWordC,aryTemplateTagsValue(5),sHPosC)
    
    If sPosD= "Before"  Then  aryTemplateTagsValue(5) =AddTextF_Before(sMyWordD,aryTemplateTagsValue(5),sHPosD)
    If sPosD="After"  Then  aryTemplateTagsValue(5) =AddTextF_After(sMyWordD,aryTemplateTagsValue(5),sHPosD)
    If sPosD="Head"  Then  aryTemplateTagsValue(5) =AddTextF_Head(sMyWordD,aryTemplateTagsValue(5),sHPosD)
    If sPosD="Foot"  Then  aryTemplateTagsValue(5) =AddTextF_Foot(sMyWordD,aryTemplateTagsValue(5),sHPosD)   
  End If
End Function

Function AddTextF_Head(AddWord, Html,HPos)   '置换关键字
   Dim objRegE,RepStr
   Set objRegE = new RegExp
   objRegE.IgnoreCase = True
   objRegE.Global = True
   objRegE.multiline = True 
   RepStr = VBCrlf 
   objRegE.Pattern= "< br /  >"
   AddWord=objRegE.Replace(AddWord, RepStr)
   Set objRegE = nothing
   
   If HPos = "Left" then AddWord = sMyLeft & AddWord & sMyLeftEnd
   If HPos = "Right" then AddWord = sMyRight & AddWord & sMyRightEnd
   AddTextF_Head = AddWord  & Html 
End Function

Function AddTextF_Before(AddWord, Html,HPos)   
   Dim objRegExp,RepStr
   Set objRegExp = new RegExp
   objRegExp.IgnoreCase = True
   objRegExp.Global = True
   objRegExp.multiline = True 
   RepStr = VBCrlf 
   objRegExp.Pattern= "< br /  >"
   AddWord=objRegExp.Replace(AddWord, RepStr)

   If HPos = "Left" then AddWord = sMyLeft & AddWord & sMyLeftEnd
   If HPos = "Right" then AddWord = sMyRight & AddWord & sMyRightEnd

   objRegExp.Pattern="<( *| */) *br( *| */)*>"
   AddTextF_Before=objRegExp.Replace(Html,  AddWord&"<br />")  
    objRegExp.Pattern="< */ *p *>"
   AddTextF_Before=objRegExp.Replace(AddTextF_Before, AddWord&"</ p>" )     
   Set objRegExp = nothing
End Function

Function AddTextF_After(AddWord, Html,HPos)  
   Dim objRegExp,RepStr
   Set objRegExp = new RegExp
   objRegExp.IgnoreCase = True
   objRegExp.Global = True
   objRegExp.multiline = True 
   RepStr = VBCrlf 
   objRegExp.Pattern= "< br /  >"
   AddWord=objRegExp.Replace(AddWord, RepStr)

   If HPos = "Left" then AddWord = sMyLeft & AddWord & sMyLeftEnd
   If HPos = "Right" then AddWord = sMyRight & AddWord & sMyRightEnd

   objRegExp.Pattern="< */ *p *>"
   AddTextF_After=objRegExp.Replace(Html, "</p>" &AddWord)  
   Set objRegExp = nothing
End Function

Function AddTextF_Foot(AddWord, Html,HPos)     '添加后置内容
   Dim objRegExp,RepStr
   Set objRegExp = new RegExp
   objRegExp.IgnoreCase = True
   objRegExp.Global = True
   objRegExp.multiline = True 
   RepStr = VBCrlf 
   objRegExp.Pattern= "< br /  >"
   AddWord=objRegExp.Replace(AddWord, RepStr)
  If HPos = "Left" then AddWord = sMySLeft & AddWord & sMySLeftEnd
  If HPos = "Right" then AddWord = sMySRight & AddWord & sMySRightEnd  
  Dim PosI,PreNum,PastNum,PosChar
  PreNum  = 0
  PastNum = 0
  If IsNumeric(sLocalNum) then PreNum = cint(sLocalNum)
  If IsNumeric(sLocalpast) then  PastNum = cint(sLocalPast)
  
  Dim HtmlLength,Top,ContentLen,Check
  HtmlLength = len(Html)  '获取日志正文长度
  PosI =HtmlLength
  ContentLen = 0
  Top = 0
 
  Do While PosI>0 And ContentLen<PreNum    
    PosChar = mid(Html,PosI,1)
    If PosChar = "<" then Top = Top+1
    If Top =0 then ContentLen = ContentLen+1      '非标签内容记长
    If PosChar = ">" then Top = Top-1
    PosI = PosI -1
  Loop
  Check = True
  If ContentLen = PreNum Then   '判断内容长度是否小于提前量
  '若内容长度等于提前量，表示可在文章中间倒计算放入 插入内容。
    Dim Matches
    '统计在后置段内符合条件的p br标签数量，为计算调整量作准备
    objRegExp.Pattern="(<( *| */) *br( *| */)*>)|(< */ *p *>)"
    Set Matches = objRegExp.Execute(right(Html, HtmlLength-PosI + 1))
    
    '根据计算的调整量，再次遍历
    PosI =HtmlLength
    ContentLen = 0
    Top = 0
    Do While (PosI>0 And ContentLen<PreNum - PastNum * Matches.count) or Top >0    
      PosChar = mid(Html,PosI,1)
      If PosChar = "<" then Top = Top+1
      If Top =0 then ContentLen = ContentLen+1      '非标签内容记长
      If PosChar = ">" then Top = Top-1
      
      PosI = PosI -1
   Loop
   '根据计算的位置，测试后段是否存在 img 等标签，
   objRegExp.Pattern="(< *img)|(< *embed)|(< */blockquote)|(< */div)|(< */span)|(< */table)"
   Check = not  objRegExp.Test(right(Html,HtmlLength-PosI + 1))          
  Else
  '小于则直接将追加内容放置在最后
    Check =False
  End If  
  IF  Check Then 
     AddTextF_Foot= left(Html,PosI+1) & AddWord & right(Html,HtmlLength-PosI -1)
  Else
     AddTextF_Foot = Html & AddWord
  End If
  Set objRegExp = nothing    
End Function



Function AddTextF_Key(ByRef Html)   '置换关键字
    '获取关键字列表,并保存在mytags 内
    Dim objRS,mytags(),tagcount,mytagslen(),i,myHtml,mytagscount(),checkok,v,tagchars,mintaglen,maxtaglen,paracount,para()
    Set objRS=objConn.Execute("SELECT [tag_ID] FROM [blog_Tag] ORDER BY [tag_Name] ASC")    
    tagcount =0
    mintaglen=256
    maxtaglen=0
    tagchars=""
    paracount =0
    If (Not objRS.bof) And (Not objRS.eof) Then
      Do While Not objRS.eof
        tagcount = tagcount+1
        ReDim preserve mytags( tagcount+1),mytagslen(tagcount+1),mytagscount(tagcount+1)
        mytags(tagcount)=Tags(objRS("tag_ID")).name   
        mytagslen(tagcount)=len(mytags(tagcount ) )
        if mytagslen(tagcount)<mintaglen then  mintaglen=mytagslen(tagcount)
        if mytagslen(tagcount)>maxtaglen then  maxtaglen=mytagslen(tagcount)
        tagchars =tagchars& mytags(tagcount)
        mytagscount(tagcount)=0
        objRS.MoveNext
      Loop    
    End IF
    objRS.Close
    Set objRS=Nothing
    '逐个文字遍历日志正文
    Dim HtmlLength,PosI,Top,ContentLen,tmpStr,ReplaceStr,PosChar,ReplaceLen,myTopHtml
    HtmlLength = len(Html)  '获取日志正文长度
    PosI =0
    Top =0     
    ContentLen = 0
    AddTextF_Key =""
    myHtml = ""
    myTopHtml=""
    tmpStr=""
    Do While PosI<=HtmlLength 
       PosI = PosI + 1
       PosChar = mid(Html,PosI,1)       
       If PosChar = "<" then 
          Top = Top+1
           ContentLen = 0          
           If myTopHtml <>"" then  
              myHtml = myHtml & myTopHtml
              myTopHtml =""
            End If            
       End If  
       If Top <>0 then tmpStr = tmpStr & PosChar
       If Top =0 then                      '确认为非标签内容
          ContentLen = ContentLen + 1      '非标签内容记长
          If tmpStr<>"" then 
             myHtml=myHtml& tmpStr
             tmpStr=""
             if len(myHtml) > 600 then 
               paracount = paracount +1
               ReDim preserve para(paracount+1)
               para(paracount) = myHtml
               myHtml=""
             end if
          End If
          myTopHtml =myTopHtml & PosChar 
          if ContentLen >= mintaglen then 
              if InStr(tagchars,PosChar)>0 then 
                  'tmpStr = tmpStr & PosChar              
                  For  i=1 to   Ubound(mytags) -1 
                     If ContentLen >= mytagslen(i)  Then 
                         If right( myTopHtml ,mytagslen(i)) = mytags(i) Then
                             ReplaceStr ="<a href="& "<#ZC_BLOG_HOST#>" &"catalog.asp?tags="  & mytags(i) &" target='_blank'>"& mytags(i)&"</a>"
                             'ReplaceStr =cint(i)&"-"&cint(mytagslen(i))&"-"&mytags(i)
                             ReplaceLen = len(myTopHtml) - mytagslen(i)
                             checkok = False 
                             mytagscount(i)= mytagscount(i) + 1
                             If (sTagInf = "0") and (mytagscount(i) = 1) then checkok = True  '只作首次
                             If sTagInf = "1" then  checkok = True '全部
                             If sTagInf ="2" or sTagInf ="3" or sTagInf = "4" Then
                               If mytagscount(i) = 1 then 
                                 checkok = True 
                               Else 
                                 v=mytagscount(i) mod Cint(sTagInf )
                                 IF v = 1 then  checkok = True 
                               End If 
                             End If                    
                             If checkok=True then 
                               myTopHtml = left(myTopHtml, ReplaceLen) &  ReplaceStr
                               ContentLen=0                                                 
                               exit for                               
                            End If
                         End If
                      End IF
                  Next 
              Else
                 if ContentLen >= maxtaglen then     
                     ContentLen = 0
                     myHtml = myHtml & myTopHtml
                     myTopHtml=""
                 End if                    
              End If
          End If
       End If 
       If PosChar = ">" then Top = Top - 1    
   Loop  
   AddTextF_Key =myHtml&myTopHtml & tmpStr
   myTopHtml=""
   tmpStr=""
   for v= 1 to paracount \2
      myTopHtml = myTopHtml & para(v)
      tmpStr = para(paracount+1 - v) & tmpStr 
    next 
    if (paracount mod 2) = 1 then 
        AddTextF_Key = myTopHtml & para(v) & tmpStr& AddTextF_Key
    else
       AddTextF_Key = myTopHtml & tmpStr&AddTextF_Key
    end if 
End Function


'安装插件
Function InstallPlugin_AddTextF()
  Call SetBlogHint(True,Empty,True)
End Function
'卸载插件
Function UnInstallPlugin_AddTextF()
  Call SetBlogHint(True,Empty,True)
End Function
%>
