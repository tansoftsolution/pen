<% on error resume next%>
<% Response.Buffer = True %> 


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Aadhaar Updated Status Report</title>
<style type="text/css">
<!--
.style8 {
	font-size: 24px;
	font-weight: bold;
        }
.style10 {color: #000099}
.style17 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; }
-->
</style>
</head>

<body>
<% 
dim totrec, cnt
Set adocon = Server.CreateObject("ADODB.Connection")
set rs=Server.CreateObject("ADODB.recordset")

db="DRIVER={Microsoft Access Driver (*.mdb)}; "
db=db & "DBQ=" & Server.MapPath("exgratia.mdb")
adocon.Open db

strSQL = "SELECT semp_nm,dob,sdesgn,emp_cd,ppono,unit,contactno,aadhaarno FROM exgra_mas where aadhaarno is null or len(trim(aadhaarno)) =0"
rs.Open strSQL, adoCon 
cnt = 1
%>
 
<div align="center" class="style10" ></div>
<p align="center"><span class="style10"><font face="Arial"><strong>TAMIL NADU CO-OPERATIVE MILK PRODUCERS' FEDERATION LTD.,: CHENNAI-35</strong></font></span><strong><font face="Arial"></font></strong></p>
<p align="center"><span class="style10"><font face="Arial"><strong>Aadhaar to be Updated Details</strong></font></span><strong><font face="Arial"></font></strong></p>
<div align="center">

 	  <table width="1000" border="1" bordercolor="#00FFFF">
        <tr>
		  <td width="10" bgcolor="#cc9900"><p class="style17">S.No</p></td>
          <td width="10" bgcolor="#cc9900"><p class="style17">Emp.Code</p></td>
          <td width="80" bgcolor="#cc9900"><p align="center" class="style17">Emp.Name</p></td>
          <td width="70" bgcolor="#cc9900"><p align="center" class="style17">Designation</p></td>
          <td width="100" bgcolor="#cc9900"><p align="center" class="style17">PPO No.</p></td>
          <td width="57" bgcolor="#cc9900"><p align="center" class="style17">Mobile No.</p></td>
          <td width="54" bgcolor="#cc9900"><p align="center" class="style17">DOB</p></td>
          <td width="54" bgcolor="#cc9900"><p align="center" class="style17">Aadhaar No.</p></td>
        </tr>
	
        <tr>
		 <% Do While not rs.EOF %>
		  <td bgcolor="#00cccc"><span class="style17">&nbsp;
                <%Response.Write((cnt))%>
          </span></td>
          <td bgcolor="#00cccc"><span class="style17">&nbsp;
                <%Response.Write(rs.fields("emp_cd"))%>
          </span></td>
          <td bgcolor="#00cccc"><span class="style17">&nbsp;
                <%Response.Write(rs.fields("semp_nm"))%>
          </span></td>
          <td bgcolor="#00cccc"><span class="style17">&nbsp;
                <%Response.Write(rs.fields("sdesgn"))%>
          </span></td>
          <td bgcolor="#00cccc"><span class="style17">&nbsp;
                <%Response.Write(rs.fields("ppono"))%>
          </span></td>
          <td bgcolor="#00cccc"><span class="style17">&nbsp;
                <%Response.Write(rs.fields("contactno"))%>
          </span></td>
          <td bgcolor="#00cccc"><span class="style17">&nbsp;
                <%Response.Write(rs.fields("dob"))%>
          </span></td>
          <td bgcolor="#00cccc"><span class="style17">&nbsp;
		        <%Response.Write(rs.fields("aadhaarno"))%>
          </span></td>
        </tr>
	           <% rs.MoveNext
			   cnt = cnt + 1
                  Loop 	%>		
      </table>

<%
  rs.close
  conn.close
  set rs=Nothing
  set conn=Nothing
%>
     <p>&nbsp;<A href="/pen/aadcount.html" >BACK</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A onclick=window.print() href="#" ;>PRINT</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>  
  <p>&nbsp;</p>


 </div>
</body>
</html>

