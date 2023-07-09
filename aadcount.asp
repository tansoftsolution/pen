
<% response.buffer=true %>
<% on error resume next%>
<%
dim totrec

Set adocon = Server.CreateObject("ADODB.Connection")

set rs=Server.CreateObject("ADODB.recordset")

db="DRIVER={Microsoft Access Driver (*.mdb)}; "
db=db & "DBQ=" & Server.MapPath("exgratia.mdb")
adocon.Open db

strSQL = "SELECT count(*) as totrec FROM exgra_mas where len(trim(aadhaarno)) > 0"
rs.Open strSQL, adoCon 
alert ("Pensioners Updated their Nos.")
Response.Write rs("totrec") & "  Pensioners Updated their Aadhaar Numbers Successfully."
rs.Close
Set rs = Nothing
Set adoCon = Nothing
%>
