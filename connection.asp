<%
Dim ConnStr
Dim DataBaseFilePath

DataBaseFilePath	= "/db/NBArticle.mdb"
ConnStr				= "Provider = Microsoft.Jet.OLEDB.4.0;Data Source =" & Server.MapPath(DataBaseFilePath)

Const sCacheName	= "NB948879421"
Const SystemFolder	= "/"
%>
