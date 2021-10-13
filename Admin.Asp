<% action = Request.QueryString("Html") %><%
 Select Case action
Case "GUANLI"  Call GUANLI()
Case "SHIBIE"  Call SHIBIE()
Case "DENGLU"  Call DENGLU()
Case "TUICHU"  Call TUICHU()

End Select
%><%if action=GUANLI then%><% N=request("N")%><%
dim xianhe
login=session("login")
if login="" then
response.redirect "?Html=DENGLU"
end if
%><%
set my_conn= Server.CreateObject("ADODB.Connection")
ConnString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("ninhu.mdb")
my_Conn.Open ConnString
%><%
edit=trim(request("edit"))
if edit="save" then
set rs=Server.CreateObject("Adodb.Recordset")
sql = "Select * from 网站 where 编号=" & request("id") & ""
rs.open Sql,my_conn,1,3
if rs.eof then
response.redirect "?ID=1&N=对不起，无用户资料，请核实您的信息！"
response.end
else if trim(request("名称"))="" or Request.Form("日期")="" or Request.Form("时间")="" then 
response.redirect "?ID=1&N=对不起，信息不完整，请返回认真填写！"
response.end
else
RS("名称")=Request.Form("名称")
RS("网址")=Request.Form("网址")
RS("状态")=Request.Form("状态")
RS("推荐")=Request.Form("推荐")
RS("点入")=Request.Form("点入")
RS("点出")=Request.Form("点出")
RS("日期")=Request.Form("日期")
RS("时间")=Request.Form("时间")
RS("登记")=Request.Form("登记")
rs.update
response.redirect "?ID=1&N=恭喜您，信息已修改，请返回刷新查看！"
Response.End
rs.close
end if
end if
else
set rs = server.CreateObject("adodb.recordset")
sql = "Select * from 网站 where 编号=" & request("id") & ""
rs.open Sql,my_conn,1,1
end if
%><title>编辑：<%=rs("名称")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<style type="text/css"><!--td, {font-family: "Palatino Linotype"; font-size: 15px}a:link{color:Black; text-decoration:none} 
a:visited{color:Black;TEXT-DECORATION: none}a:hover{color:red;text-decoration:underline} --></style>
<TABLE width=100% height=100% align=center border=0><TR height=320 align=center>
<TD align=center background=http://www.ninhu.com/cangku/guang.gif style="background-position:center">
<TABLE width=1200 align=center background=http://www.ninhu.com/2010-2020/1070.PNG border=0>
<!--编辑开始--><TR align=center>
<FORM name=searchForm method=get action=?ID= target=_self>
<TD><input disabled type=button value=操作提示： style="width:100%;height:35px"></td>
<td colSpan=2><input onfocus="this.value='';this.focus()" value='<%=N%>' style="width:100%;height:35px;text-align:center" name=ID></td></FORM>
<FORM action=?ID=<%=request("id")%> method=post name=reg target=_self>
<TD><input disabled type=button value=累积来访： style="width:100%;height:35px"></td>
<td colSpan=2><input maxlength=10 name="点入" value='<%=rs("点入")%>' style="width:100%;height:35px;text-align:center"></td>
</TR><TR align=center>
<td><input disabled type=button value=网站名称： style="width:100%;height:35px"></td>
<td colSpan=2><input maxlength=50 name="名称" value='<%=rs("名称")%>' style="width:100%;height:35px;text-align:center"></td>
<td><input disabled type=button value=来访时间： style="width:100%;height:35px"></td>
<td colSpan=2><input maxlength=19 name="日期" value='<%=rs("日期")%>' style="width:100%;height:35px;text-align:center"></td>
</TR><TR align=center>
<td><input disabled type=button value=网站地址： style="width:100%;height:35px"></td>
<td colSpan=2><input maxlength=50 name="网址" value='<%=rs("网址")%>' style="width:100%;height:35px;text-align:center"></td>
<TD><input disabled type=button value=累积点出： style="width:100%;height:35px"></td>
<td colSpan=2><input maxlength=10 name="点出" value='<%=rs("点出")%>' style="width:100%;height:35px;text-align:center"></td>
</TR><TR align=center>
<td><input disabled type=button value=登记时间： style="width:100%;height:35px"></td>
<td colSpan=2><input maxlength=19 name="登记" value='<%=rs("登记")%>' style="width:100%;height:35px;text-align:center"></td>
<td><input disabled type=button value=点出时间： style="width:100%;height:35px"></td>
<td colSpan=2><input maxlength=19 name="时间" value='<%=rs("时间")%>' style="width:100%;height:35px;text-align:center"></td>
</TR><TR align=center>
<td><input disabled type=button value=当前时间： style="width:100%;height:35px"></td>
<td colSpan=2><input disabled value='<%=now%>' style="width:100%;height:35px;text-align:center"></td>
<td><input disabled type=button value=网站状态： style="width:100%;height:35px"></td>
<td colSpan=2><select style="width:20%;height:35px;text-align:center" name="状态"><option value="<%=rs("状态")%>"><%=rs("状态")%></option>
<option value="0">正常</option><option value="1">锁定</option></select><select style="width:20%;height:35px;text-align:center" name="推荐"><option value="<%=rs("推荐")%>"><%=rs("推荐")%></option><option value="0">正常</option><option value="1">链接</option></select><input type="hidden" name="edit" value="save"><input type=submit value=' 确认修改 ' name="B1" style='width:40%;height:35px;text-align:center'><input onClick="location.href='?Html=TUICHU'" type=button value="退出" style="width:20%;height:35;text-align:center"/></td></TR></FORM><!--编辑结束-->
<TR align=center><TD width=16%><TABLE width=100% align=center border=1>
<TR><td width=100% height=40 align=center>【最新来访站点】</TD></TR><TR><%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select top 100 * from 网站 where 状态=0 order by 日期 desc"
'response.write sql
rs.open sql,my_conn,1,1
 %><%
  if rs.eof and rs.bof then
 response.write "<TD width=100% height=368 align=center>没有数据</TD>"
 else
 filename=1
%><%do while not rs.eof and filename<200%><TD width=100% height=35 align=center>
<A target=_self href='?ID=<%=rs("编号")%>&N=<% =replace(rs("网址"),"http://","") %>'><%=left(rs("名称"),7)%></A></TD><% 
if filename mod 1 =0 then
Response.Write("</TR>")
end if   
rs.movenext
filename=filename+1
if filename>10 then Exit Do
loop
end if
rs.Close
set rs=nothing
%></TABLE></TD><TD width=16%><TABLE width=100% align=center border=1>
<TR><td width=100% height=40 align=center>【最久来访站点】</TD></TR><TR><%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select top 100 * from 网站 where 状态=0 order by 日期 asc"
'response.write sql
rs.open sql,my_conn,1,1
 %><%
  if rs.eof and rs.bof then
 response.write "<TD width=100% height=368 align=center>没有数据</TD>"
 else
 filename=1
%><%do while not rs.eof and filename<200%><TD width=100% height=35 align=center>
<A target=_self href='?ID=<%=rs("编号")%>&N=<% =replace(rs("网址"),"http://","") %>'><%=left(rs("名称"),7)%></A></TD><% 
if filename mod 1 =0 then
Response.Write("</TR>")
end if   
rs.movenext
filename=filename+1
if filename>10 then Exit Do
loop
end if
rs.Close
set rs=nothing
%></TABLE></TD><TD width=16%><TABLE width=100% align=center border=1>
<TR><td width=100% height=40 align=center>【来访排行站点】</TD></TR><TR><%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select top 100 * from 网站 where 状态=0 order by 点入 desc"
'response.write sql
rs.open sql,my_conn,1,1
 %><%
  if rs.eof and rs.bof then
 response.write "<TD width=100% height=368 align=center>没有数据</TD>"
 else
 filename=1
%><%do while not rs.eof and filename<200%><TD width=100% height=35 align=center>
<A target=_self href='?ID=<%=rs("编号")%>&N=<% =replace(rs("网址"),"http://","") %>'><%=left(rs("名称"),7)%></A></TD><% 
if filename mod 1 =0 then
Response.Write("</TR>")
end if   
rs.movenext
filename=filename+1
if filename>10 then Exit Do
loop
end if
rs.Close
set rs=nothing
%></TABLE></TD><TD width=16%><TABLE width=100% align=center border=1>
<TR><td width=100% height=40 align=center>【点出排行站点】</TD></TR><TR><%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select top 100 * from 网站 where 状态=0 order by 点出 desc"
'response.write sql
rs.open sql,my_conn,1,1
 %><%
  if rs.eof and rs.bof then
 response.write "<TD width=100% height=368 align=center>没有数据</TD>"
 else
 filename=1
%><%do while not rs.eof and filename<200%><TD width=100% height=35 align=center>
<A target=_self href='?ID=<%=rs("编号")%>&N=<% =replace(rs("网址"),"http://","") %>'><%=left(rs("名称"),7)%></A></TD><% 
if filename mod 1 =0 then
Response.Write("</TR>")
end if   
rs.movenext
filename=filename+1
if filename>10 then Exit Do
loop
end if
rs.Close
set rs=nothing
%></TABLE></TD><TD width=16%><TABLE width=100% align=center border=1>
<TR><td width=100% height=40 align=center>【友情链接站点】</TD></TR><TR><%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select top 100 * from 网站 where 推荐=1 order by 日期 desc"
'response.write sql
rs.open sql,my_conn,1,1
 %><%
  if rs.eof and rs.bof then
 response.write "<TD width=100% height=368 align=center>没有数据</TD>"
 else
 filename=1
%><%do while not rs.eof and filename<200%><TD width=100% height=35 align=center>
<A target=_self href='?ID=<%=rs("编号")%>&N=<% =replace(rs("网址"),"http://","") %>'><%=left(rs("名称"),7)%></A></TD><% 
if filename mod 1 =0 then
Response.Write("</TR>")
end if   
rs.movenext
filename=filename+1
if filename>10 then Exit Do
loop
end if
rs.Close
set rs=nothing
%></TABLE></TD><TD width=16%><TABLE width=100% align=center border=1>
<TR><td width=100% height=40 align=center>【最新加入站点】</TD></TR><TR><%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select top 100 * from 网站 order by 编号 desc"
'response.write sql
rs.open sql,my_conn,1,1
 %><%
  if rs.eof and rs.bof then
 response.write "<TD width=100% height=368 align=center>没有数据</TD>"
 else
 filename=1
%><%do while not rs.eof and filename<200%><TD width=100% height=35 align=center>
<A target=_self href='?ID=<%=rs("编号")%>&N=<% =replace(rs("网址"),"http://","") %>'><%=left(rs("名称"),10)%></A></TD><% 
if filename mod 1 =0 then
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

<TR><TD width=16% colSpan=6><TABLE width=100% align=center border=1><TR><%
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select * from 网站 where 状态=1 order by 日期 desc"
'response.write sql
rs.open sql,my_conn,1,1
 %><%
  if rs.eof and rs.bof then
 response.write "<TD colSpan=6 width=100% height=35 align=center>没有数据</TD>"
 else
 filename=1
%><%do while not rs.eof and filename<200%><TD width=15% height=35 align=center>
<A target=_self href='?ID=<%=rs("编号")%>&N=<% =replace(rs("网址"),"http://","") %>'><S><%=left(rs("名称"),8)%></S></A></TD><% 
if filename mod 6 =0 then
Response.Write("</TR>")
end if   
rs.movenext
filename=filename+1
if filename>30 then Exit Do
loop
end if
rs.Close
set rs=nothing
%></TABLE></TD></TR></TABLE></TD></TR></TABLE><% End IF %>


<% Sub SHIBIE() %><%
pass=trim(request.form("password")) 
if pass="" then
response.redirect "?Html=DENGLU&N=账号或密码不能为空"
end if
if pass="admin" then
session("login")="yes"
response.redirect "?ID=1&N=恭喜您，登录成功。现在可以对网站进行管理了！"
else
response.redirect "?Html=DENGLU&N=账号和密码验证错误"
end if
%><% End Sub %>


<% Sub DENGLU() %><title>后台登录</title><% N=request("N")%>
<TABLE width=100% height=100% align=center border=0><TR height=320 align=center><TD align=center background=http://www.ninhu.com/cangku/guang.gif style="background-position:center">
<TABLE height=100 align=center border=0><FORM action=?Html=SHIBIE method=post target=_self><TR align=center>
<TD><input placeholder="账号" maxlength=25 style='width:300;height:35;text-align:center' value='<%=N%>'></TD>
<TD rowSpan=2><input type="submit" name="Submit" value="登录" style="width:100px;height:85px;font-size:25px"></TR>
<TR align=center><TD><input placeholder="密码" maxlength=15 style='width:300;height:35;text-align:center' value='' type=password name=password></TD></TR></FORM>
</TD></TR></TABLE></TR></TABLE>
<% End Sub %>


<% Sub TUICHU() %><% Session.Abandon()
Response.Redirect "/"%><% End Sub %>