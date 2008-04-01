<%
Dim ConnStr
Dim DataBaseFilePath

DataBaseFilePath	= "/db/mobile.mdb"
ConnStr				= "Provider = Microsoft.Jet.OLEDB.4.0;Data Source =" & Server.MapPath(DataBaseFilePath)

Const sCacheName	= "NB511654079"
Const SystemFolder	= "/"
%>
