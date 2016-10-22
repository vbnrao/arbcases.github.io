<%
'#######################################
'This file points to the database and 
'opens a connection to allow
'querying of the data held
'Change the location of the Server.MapPath
'line if you moved the database
'#######################################

Dim conn, ConnectString
ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("ArbCMS_Data.mdb") & ";Persist Security Info=False"
Set conn = Server.CreateObject("ADODB.Connection")
conn.open ConnectString
%>
