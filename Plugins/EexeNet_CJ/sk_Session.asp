  <%
  '===================================
 '是不是管理员登录了，不是的话退出去了
'文件修改于2007-12-30 by http://eexe.net
'================================================
 
  if session(CookieName&"_System")<>true then 
        Response.write("<script language='javascript'>alert('你还没有登录后台呢');location.href='http://eexe.net'</script>")
   	  '  Response.Redirect("http://www.eexe.net")
  End if
  
  
  %>