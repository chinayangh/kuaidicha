<%OPTION EXPLICIT%>
<%Server.ScriptTimeOut=5000%>
<!--#include FILE="upload_5xsoft.inc"-->
<%
If IsEmpty(Session("username")) Then
	response.write "<script>window.location.href='index.html'</script>"
else
dim upload,file,formName,formPath,fs,fs1
set upload=new upload_5xsoft ''�����ϴ�����
Set fs=Server.CreateObject("Scripting.FileSystemObject")
'Set fs1=Server.CreateObject("Scripting.FileSystemObject")

formPath="file/"'·��

for each formName in upload.objFile ''�г������ϴ��˵��ļ�
	set file=upload.file(formName)  ''����һ���ļ�����
	if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
		file.SaveAs Server.mappath(formPath&file.FileName)   ''�����ļ�
		fs.CopyFile "c:\inetpub\kuaidi\kdinfo.mdb","c:\inetpub\kuaidi\bak\kdinfo.mdb."&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now),True
		fs.CopyFile "c:\inetpub\kuaidi\file\"&file.FileName,"c:\inetpub\kuaidi\file\"&file.FileName&Session("username"),True 
		fs.CopyFile "c:\inetpub\kuaidi\file\"&file.FileName,"c:\inetpub\kuaidi\file\book2.xls",True    '��fs��CopyFile���������ļ�

	end if
	set file=nothing
next

set upload=nothing  ''ɾ���˶���
end if
%>
<script>window.parent.Finish("�ϴ��ļ��ɹ���");</script>