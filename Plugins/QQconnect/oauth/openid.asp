<script Language="JScript" runat="server">
function QQConnection(strPath){
  try{
	ConnQQ = Server.CreateObject("ADODB.Connection");
	ConnQQ.connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Server.MapPath(strPath);
	ConnQQ.open();
  }catch(e){
	Response.Write('<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" /><div style="font-size:12px;font-weight:bold;border:1px solid #006;padding:6px;background:#fcc">数据库连接出错，请检查连接字串!</div>');
	Response.End
  }
}
</script>
<%
Call QQConnection("Plugins/QQconnect/oauth/#54bq_openid.asp")
%>