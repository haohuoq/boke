<!--#include file="../sysfile/config.asp" -->
<%
	Dim db, myconn,conn
		db="../"&db&""
		myconn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
	On Error Resume Next
		Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open myconn

	If Err Then
		err.Clear
		Set Conn = Nothing
		response.write"���ݿ����ӳ������������ִ���"
'		response.write db
		Response.End
	End If
sub CloseConn()
conn.close
set conn=nothing
end sub
Const SafePass="1,1,4,2,3"
'���ã�����ͨ���������ݿ��SQLע��õ��˹���Ա������������Բ��ܽ���ϵͳ
'
'��1λ	�Ƿ����ð�ȫ���� Ϊ0ʱ������ Ϊ1ʱ����
'��2λ	ȡ��֤���еĵڼ�λ�������㣬ȡ1-4֮�������
'��3λ	ȡ��֤���еĵڼ�λ�������㣬ȡ1-4֮�������
'��4λ	��ȡ�õ���λ��֤����ʲô���㣬1Ϊ�ӷ����㣻2Ϊ�˷�����
'��5λ	���õ��Ľ�����뵽����ĵڼ�λ����
'
'���簲ȫ���������Ϊ1��1��3��2��5  ��Ϊ���ð�ȫ�룬����֤��ĵ�һλ�͵���λ��˵Ľ�����뵽����ĵ���λ����
'������½ʱ ��������֤��Ϊ3568 ����Ա����ΪTryLogin
'����Ӧ�����������ΪTryLo18gin
%>