<%
const NHURL="www.cnzit.com" '	您网站的地址
const NHNAME="站长推广网" '	您网站的名称
const NINHU="www.ninhu.com" '	版权地址，请勿修改
%><body background=http://<%=NINHU%>/2020-2030/0145.PNG style="background-attachment:fixed">
<style># html { filter:progid:DXImageTransform.Microsoft.BasicImage(grayscale=1); }
body,td,th {font-size:15px;color: #333333;text-decoration:none;line-height:40px}
body {margin: 0px;background-color: #FDFDFD; font-family:Arial;}
a:link { color: #333333;text-decoration: none;}
a:visited { color: #333333;text-decoration: none;}
a:hover {color:009933;text-decoration: none;}
nav{width:100%;margin:0em auto;}
ul{margin:0px;padding:0px;list-style:none;}
ul.dropdown li{font-weight:bold;float:left;width:100%;position:relative;background:#444444;}
ul.dropdown li a{display:block;padding:0px 0px;color:#FAF5F3;position:relative;text-align:center;text-decoration:none;}
ul.dropdown li a:hover,
ul.dropdown li a.hover{background:#FAF5F3;position:relative;color:#444444;}
.DIVA{text-align:center;border:1px dashed #000000;padding:0px;}
.DIVB{text-align:center;border:1px dashed #FF6600;padding:0px;}
.DIVC{text-align:center;border:1px dashed #FF00FF;padding:0px;}
.DIVD{text-align:center;border:1px dashed #008000;padding:0px;}
.DIVE{text-align:center;border:1px dashed #0000FF;padding:0px;}
.DIVF{text-align:center;border:1px dashed #993366;padding:0px;}
</style><%if Request("N")="" then %><% 
dim db,conn,connstr
db=db1 & "ninhu.mdb" '数据库路径
set Conn = server.CreateObject("ADODB.Connection")
connstr="provider=microsoft.jet.oledb.4.0;data source=" & server.MapPath(db)
conn.Open connstr

IF Session("addurl")<>"" Then
SQL="Select * From 网站 Where 网址='"& Session("addurl") &"'"
Set RS=Server.CreateObject("ADODB.Recordset")
RS.Open SQL,Connstr,3,2
IF Not RS.Eof Then
RS("时间")=Now
RS.Update
End IF
End IF
%><%
Action="reg"
IF Request.Form("SubmitReg")="OK" Then
URL=Request.Form("URL")
TITLE=Request.Form("TITLE")
RegErr=""
RegOK=True
IF Len(TITLE)<2 or Len(TITLE)>10 or Len(URL)<12 or Len(URL)>50 Then
RegErr=RegErr & "<!--WWW.NINHU.COM-->"
RegOK=False
Action="error"
End IF
IF RegErr="" Then
SQL="Select * From 网站 Where 网址='"& URL &"'"
Set RS=Server.CreateObject("ADODB.Recordset")
RS.Open SQL,Connstr,3,2
IF RS.EOF Then
RS.AddNew
RS("网址")=URL
RS("名称")=TITLE
RS("登记")=Now
RS("点入")=0
RS("点出")=0
RS("推荐")=0
RS("状态")=0
RS.Update
Action="suc"
RegOK=True
Else
RegErr=RegErr & "恭喜您，提交成功！做好我站链接，来访次数越多，展现的机会就越多！<BR>网站名称：站长推广网，链接地址：http://www.cnzit.com，LOGO地址：http://www.ninhu.com/cangku/link.gif<HR>"
RegOK=False
Action="error"
End IF
End IF
End IF
%><html><head>
<title><%=NHNAME%>（V1.5）免费收录网站，迅速提升流量！ </title>
<meta name="keywords" content="自助链,自动链,<%=NHNAME%>">
<meta name="Description" content="<%=NHNAME%>是一个非常专业的自助链网站，免费为新老站点提供永久性的自动链广告位，让新老站点获得更多的被点击机会。无需注册，无需审核，提交网站，做上链接，清除缓存，点击一次，完成收录！">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="shortcut icon" href="http://<%=NINHU%>/2020-2030/0141.PNG" />
<base target=_blank><% N=request("N")%></HEAD><%
set my_conn= Server.CreateObject("ADODB.Connection")
ConnString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("ninhu.mdb")
my_Conn.Open ConnString
  if Request.ServerVariables("HTTP_REFERER")<>off then
  url   =   Request.ServerVariables("HTTP_REFERER")
  if url<>"" then 
  domain=split(url,"/")(2)  
  b=ubound(split(domain,"."))
  if b=1 then
  domain="http://www."&domain
  else
  domain="http://"&domain
  end if
if domain<>"http://www.cnzit.com" then
weburl=domain
set rrs=server.CreateObject("adodb.recordset")
rsql="select * from 网站 where 网址='"&weburl&"'"
rrs.open rsql,my_conn,1,3
if rrs.eof and rrs.bof then
else
rrs("日期")=now()
rrs("点入")=rrs("点入")+1
rrs.update
end if
end if
end if
end if
%></head>
<TABLE width=100% height=100% border=0 align=center><TR><TD align=center>
<TABLE width=1200 height=80 border=0 align=center><FORM action=http://www.baidu.com/s target=_self><TR>
<TD rowSpan=2 colSpan=2 align=left><A href=http://<%=Request.ServerVariables("Http_Host")%><%=Request.ServerVariables("Url")%> target=_self><img alt=<%=NHNAME%> src=http://<%=NINHU%>/cangku/tuiguang.gif border=0></A></TD>
<TD colSpan=2 align=right style="padding-right:5"><A href=http://www.baidu.com/baidu?word=<%=NHNAME%>><img alt=百度一下 src=http://<%=NINHU%>/cangku/baidu.png border=0></A></TD>
<TD colSpan=3><INPUT id=query maxlength=50 style="width:100%;height:35px;padding-left:5" value='<%=S%><%=N%>' name=wd></TD>
<TD><INPUT type=button value='百度一下' onClick=window.open('https://www.baidu.com/baidu?word='+form.query.value) style='width:100%;height:35px;'></TD></TR><TR>
<TD colSpan=2><INPUT style="width:100%;height:35px;color:#FDFDFD;background-image:url(http://www.ninhu.com/2020-2030/1044.GIF)" type=button value="欢迎访问：<%=NHNAME%>" onclick="window.location.reload()"></TD>
<TD><INPUT type=button value='好搜' onClick=window.open('https://www.so.com/s?q='+form.query.value) style='width:100%;height:35px;'></TD>
<TD><INPUT type=button value='搜狗' onClick=window.open('https://www.sogou.com/web?query='+form.query.value) style='width:100%;height:35px;'></TD>
<TD><INPUT type=button value='多吉' onClick=window.open('https://www.dogedoge.com/results?q='+form.query.value) style='width:100%;height:35'></TD>
<TD><INPUT type=button value='贰地' onClick=window.open('https://www.2dil.com/ZNSS-1?S='+form.query.value) style='width:100%;height:35px;'></TD>
</TR></FORM><NAV><TR><%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select * from 网站 where 状态=0 order by 日期 desc"
'response.write sql
rs.open sql,my_conn,1,1
 %><%
  if rs.eof and rs.bof then
 response.write "<TD width=100% height=300 colSpan=2 align=center>没有数据</TD>"
 else
 filename=1
%><%do while not rs.eof and filename<200%>
<TD width=12% height=35 CLASS=DIVB align=center><UL class=dropdown><LI><A href='?ID=<%=rs("编号")%>&N=DC' title="网址：<% =replace(rs("网址"),"http://","") %>&#13来访：<%=rs("点入")%>次&#13日期：<%=rs("日期")%>&#13点出：<%=rs("点出")%>次&#13日期：<%=rs("时间")%>"><%=left(rs("名称"),8)%></A></LI></UL></TD><% 
if filename mod 8 =0 then
Response.Write("</TR>")
end if   
rs.movenext
filename=filename+1
if filename>40 then Exit Do
loop
end if
rs.Close
set rs=nothing
%></NAV>
<%If Action="reg" Then%><FORM method=post action="http://<%=Request.ServerVariables("Http_Host")%><%=Request.ServerVariables("Url")%>" target=_self>
<TR><TD colSpan=2><input disabled type=button value="加入<%=NHNAME%>，免费推广您的网站：" style="width:100%;height:40;font-size:15px"></TD>
<TD colSpan=2><input maxlength=8 name="TITLE" placeholder="网站名称，八个字内" type=text style="width:100%;height:40px;padding-left:5"></TD>
<TD colSpan=3><input maxlength=50 value=http://www. placeholder="网站网址，以 http:// 开头 后面不要带 /" name=URL value='' style="width:100%;height:40px;padding-left:5"></TD>
<INPUT type=hidden name="reg" value="OK"><INPUT type=hidden name="submitreg" value="OK">
<TD><input type="submit" name="Submit" value="快速提交" style="width:100%;height:40px;color:#FF6600"></TD></TR></FORM>
<%ElseIF Action="error" Then%><FORM method=post action="http://<%=Request.ServerVariables("Http_Host")%><%=Request.ServerVariables("Url")%>" target=_self>
<TR><TD colSpan=2><input disabled type=button value="网站名称或网址请按要求填写，重填：" style="width:100%;height:40;font-size:15px;color:#FF0000"></TD>
<TD colSpan=2><input maxlength=8 name="TITLE" placeholder="网站名称，两个字以上，八个字以内" type=text style="width:100%;height:40px;padding-left:5"></TD>
<TD colSpan=3><input maxlength=50 value=http://www. placeholder="网站网址，以 http:// 开头 后面不要带 /" name=URL value='' style="width:100%;height:40px;padding-left:5"></TD>
<INPUT type=hidden name="reg" value="OK"><INPUT type=hidden name="submitreg" value="OK">
<TD><input type="submit" name="Submit" value="重新提交" style="width:100%;height:40px;color:#FF0000"></TD></TR></FORM>
<%Response.Write RegErr%><%ElseIf Action="suc" Then%>
<TR><TD colSpan=2><input disabled type=button value="恭喜您，提交成功，请获取连接代码：" style="width:100%;height:40;font-size:15px;color:#0000FF"></TD>
<TD colSpan=5><input readonly value="网站名称：<%=NHNAME%>，链接地址：http://<%=NHURL%>，LOGO地址：http://<%=NINHU%>/cangku/link.gif" style='width:100%;height:40px;text-align:center;'></TD>
<TD><input type="submit" disabled value="提交成功" style="width:100%;height:40px;color:#0000FF"></TD></TR><%Else
Response.Redirect "index.asp"
End IF%><TR align=center>
<td colSpan=2><input disabled type=button value="【活跃（来访）排行榜】" style="width:100%;height:40;font-size:16px"></TD>
<td colSpan=2><input disabled type=button value="【热度（点出）排行榜】" style="width:100%;height:40;font-size:16px"></TD>
<td colSpan=2><input disabled type=button value="【最新热度（点出）榜】" style="width:100%;height:40;font-size:16px"></TD>
<td colSpan=2><input disabled type=button value="【最新收录（新站）榜】" style="width:100%;height:40;font-size:16px"></TD></TR>
<TR><TD width=25% colSpan=2><TABLE width=100% align=center border=none><TR><%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select top 100 * from 网站 where 状态=0 order by 点入 desc"
'response.write sql
rs.open sql,my_conn,1,1
 %><%
  if rs.eof and rs.bof then
 response.write "<TD width=100% height=300 colSpan=2 align=center>没有数据</TD>"
 else
 filename=1
%><%do while not rs.eof and filename<200%>
<TD width=50% height=35 CLASS=DIVC align=center><A href='?ID=<%=rs("编号")%>&N=DC' title="网址：<% =replace(rs("网址"),"http://","") %>&#13来访：<%=rs("点入")%>次&#13日期：<%=rs("日期")%>"><%=left(rs("名称"),8)%></A></TD><% 
if filename mod 2 =0 then
Response.Write("</TR>")
end if   
rs.movenext
filename=filename+1
if filename>10 then Exit Do
loop
end if
rs.Close
set rs=nothing
%></TABLE></TD>
<TD width=25% colSpan=2><TABLE width=100% align=center border=none><TR><%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select top 100 * from 网站 where 状态=0 order by 点出 desc"
'response.write sql
rs.open sql,my_conn,1,1
 %><%
  if rs.eof and rs.bof then
 response.write "<TD width=100% height=300 colSpan=2 align=center>没有数据</TD>"
 else
 filename=1
%><%do while not rs.eof and filename<200%>
<TD width=50% height=35 CLASS=DIVD align=center><A href='?ID=<%=rs("编号")%>&N=DC' title="网址：<% =replace(rs("网址"),"http://","") %>&#13点出：<%=rs("点出")%>次&#13日期：<%=rs("时间")%>"><%=left(rs("名称"),8)%></A></TD><% 
if filename mod 2 =0 then
Response.Write("</TR>")
end if   
rs.movenext
filename=filename+1
if filename>10 then Exit Do
loop
end if
rs.Close
set rs=nothing
%></TABLE></TD>
<TD width=25% colSpan=2><TABLE width=100% align=center border=none><TR><%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select top 100 * from 网站 where 状态=0 order by 时间 desc"
'response.write sql
rs.open sql,my_conn,1,1
 %><%
  if rs.eof and rs.bof then
 response.write "<TD width=100% height=300 colSpan=2 align=center>没有数据</TD>"
 else
 filename=1
%><%do while not rs.eof and filename<200%>
<TD width=50% height=35 CLASS=DIVE align=center><A href='?ID=<%=rs("编号")%>&N=DC' title="网址：<% =replace(rs("网址"),"http://","") %>&#13点出：<%=rs("点出")%>次&#13日期：<%=rs("时间")%>"><%=left(rs("名称"),8)%></A></TD><% 
if filename mod 2 =0 then
Response.Write("</TR>")
end if   
rs.movenext
filename=filename+1
if filename>10 then Exit Do
loop
end if
rs.Close
set rs=nothing
%></TABLE></TD>
<TD width=25% colSpan=2><TABLE width=100% align=center border=none><TR><%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select top 100 * from 网站 where 状态=0 order by 编号 desc"
'response.write sql
rs.open sql,my_conn,1,1
 %><%
  if rs.eof and rs.bof then
 response.write "<TD width=100% height=300 colSpan=2 align=center>没有数据</TD>"
 else
 filename=1
%><%do while not rs.eof and filename<200%>
<TD width=50% height=35 CLASS=DIVF align=center><A href='?ID=<%=rs("编号")%>&N=DC' title="网址：<% =replace(rs("网址"),"http://","") %>&#13编号：<%=rs("编号")%>号&#13日期：<%=rs("登记")%>"><%=left(rs("名称"),8)%></A></TD><% 
if filename mod 2 =0 then
Response.Write("</TR>")
end if   
rs.movenext
filename=filename+1
if filename>10 then Exit Do
loop
end if
rs.Close
set rs=nothing
%></TABLE></TD></TR>
<NAV><TR><%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select top 100 * from 网站 where 推荐=1 order by 日期 desc"
'response.write sql
rs.open sql,my_conn,1,1
 %><%
  if rs.eof and rs.bof then
 response.write "<TD width=100% height=35 colSpan=2 align=center>没有数据</TD>"
 else
 filename=1
%><%do while not rs.eof and filename<200%>
<TD width=12% height=35 CLASS=DIVA align=center><UL class=dropdown><LI><A href='<%=rs("网址")%>'><%=left(rs("名称"),8)%></A></LI></UL></TD><% 
if filename mod 8 =0 then
Response.Write("</TR>")
end if   
rs.movenext
filename=filename+1
if filename>8 then Exit Do
loop
end if
rs.Close
set rs=nothing
%></NAV></TABLE><%end if%><%if Request("N")="DC" then %><%
set my_conn= Server.CreateObject("ADODB.Connection")
ConnString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("ninhu.mdb")
my_Conn.Open ConnString
%><%
Response.Expires=0
if not isempty(request.querystring("id")) then
id=request.querystring("id")
set rs=server.createobject("adodb.recordset")
sql="select * from 网站 where 编号="&ID
rs.open sql,my_conn,1,3
If Request.Cookies("RateSite")("Site" & intSiteID)<> "True" then
Response.Cookies("RateSite")("Site" & intSiteID) = "True"
Response.Cookies("RateSite").Expires = Now() + 1
rs("点出")=rs("点出")+1
end if
rs("时间")=now()
rs.update
rs.Close
set rs=nothing
end if
%><%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select top 100 * from 网站 where 编号="&ID
'response.write sql
rs.open sql,my_conn,1,1
 %><script type="text/javascript">
location.href = "http://<% =replace(rs("网址"),"http://","") %>";
</script><%end if%><center><A href=http://<%=NHURL%>><%=NHNAME%></A>已安全运行了 13 年 <%=Month(now())%> 个月 <%=day(now())%> 天 <%=hour(now())%> 小时 <%=Minute(now())%> 分 <%=Second(now())%> 秒，累积IP：10,000,000+，累积PV：100,000,000+<BR>版权所有：<A href="http://xxx.2dil.com" target="_blank">站长推广网</A>　蜀ICP备2020035862号　技术支持：<A href=http://<%=NINHU%>>您湖网络</A><BR>
<script src=http://www.ninhu.com/Java.Asp?ID=6151></script> |
<script src=http://www.ninhu.com/Java.Asp?ID=6152></script> |
<script src=http://www.ninhu.com/Java.Asp?ID=6153></script> |
<script src=http://www.ninhu.com/Java.Asp?ID=6154></script> |
<script src=http://www.ninhu.com/Java.Asp?ID=6155></script></center>
</TD></TR></TABLE><img width=0 height=0 src="//ia.51.la/go1?id=3307088&pvFlag=1" style="border:none" /></BODY></HTML>