
<% response.buffer=true %>
<% on error resume next%>
<%
'response.write(request.querystring("mobno"))
'Response.Write ("<br>")
'response.write(request.querystring("aadno"))
dim con,wmobno,waadno,sdocon,rs,db,rs1
wmobno = request.form("mobno")
waadno = request.form("aadno")
'response.write(wmobno)
'Response.Write("<br>")
'response.write(waadno)
'Response.Write ("<br>")
Set adocon = Server.CreateObject("ADODB.Connection")

set rs=Server.CreateObject("ADODB.recordset")

db="DRIVER={Microsoft Access Driver (*.mdb)}; "
db=db & "DBQ=" & Server.MapPath("exgratia.mdb")
adocon.Open db


'Initialise the strSQL variable with an SQL statement to query the databas

strSQL = "SELECT semp_nm,dob,sdesgn,emp_cd,ppono,unit,contactno,aadhaarno FROM exgra_mas where contactno = '" & Request.Form("mobno") &"' and dob = '" & Request.Form("dob") &"' "
rs.Open strSQL, adoCon 
'Response.Write ("<br>")
'Response.Write "Employee Code : " &(rs("emp_cd"))
 
if rs.EOF and rs.BOF then
     response.write "You Mobile No. or Date of Birth not correct. Try again."
else
       Response.Write ("<br>")
       Response.Write "Employee Code : " &(rs("emp_cd"))
       Response.Write ("<br>")
       Response.Write "Employee Name : " &(rs("semp_nm"))
       Response.Write ("<br>")
       Response.Write "Date of Birth :  " &(rs("dob"))
       Response.Write ("<br>")
       Response.Write "Designation   : " &(rs("sdesgn"))
       Response.Write ("<br>")
       Response.Write "PPO No.       : " &(rs("ppono"))
       Response.Write ("<br>")
	   Response.Write "Contact No.   : " &(rs("contactno"))
	   Response.Write ("<br>")
       Response.Write "Unit Name     : " &(rs("unit"))
       Response.Write ("<br>")
      
  	   Response.Write "Aadhaar No    : " &(rs("aadhaarno"))
       Response.Write ("<br>")

 	   Response.Write "Aadhaar No New:" 
       Response.Write(waadno)
       Response.Write ("<br>")
      
	   strSQL =  "Update exgra_mas set aadhaarno = '" & Request.Form("aadno") &"' where contactno = '" & Request.Form("mobno") &"' and dob = '" & Request.Form("dob") &"' "
	   set rs1=adocon.Execute(strSQL)
       Response.Write " Your Aadhaar No. Updated Successfully."
end if

rs.Close
Set rs = Nothing
Set adoCon = Nothing
%>




