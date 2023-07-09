<% on error resume next%>

<HTML>
<HEAD>
<TITLE>JPLC - AADHAAR UPDATE SCREEN</TITLE>
<script language="javascript">

function chk()
{
	frm=document.loginfrm;
	if(frm.mobno.value=="")	
	{
		alert(" Enter Correct mOBILR nO");
		frm.mobno.focus;
		return false;
	}

	if(frm.dob.value=="")	
	{
		alert(" Enter Correct Date of Birth");
		frm.dob.focus;
		return false;
	}
	
}
</script>
</HEAD>
<BODY BGCOLOR="#FFFFFF">
<p align=center><font face="arial" size=4 color=blue><span class="style10"><font face="Arial"><strong>AAVIN EX-GRATIA PENSION 
  ONLINE JEEVAN PRAMAAN LIFE CERIFICATE AADHAAR NO ENTRY SYSTEM</strong></font></span><strong><font face="Arial">.</font></strong></font></p>
<p align=center><font face="arial" size=4 color=blue> Pensioners login for Update Aadhaar No. </font></p>
<FORM METHOD=POST ACTION="../pen/aadharpost.asp" Name="loginfrm">
  <TABLE border="0" cellpadding="1" cellspacing="1" align="center" width="36%" bordercolor="#008080">
    <tr  bgcolor="#006666"> 
      <td colspan=2>&nbsp;</td>
    </tr>
   
    <TR bgcolor="#E5E5E5"> 
      <TD width="73%"><font color="darkblue" size="2" face="arial">MOBILE No.</font></TD>
      <TD width="27%"><INPUT NAME="mobno" TYPE="text" maxlength="10" size="10"></TD>
    </TR>
    <TR bgcolor="#E5E5E5"> 
      <TD><font color="darkblue" size="2" face="arial">Date of Birth(DD/MM/YYYY)</font></TD>
      <TD><INPUT NAME="dob" TYPE="text" maxlength="10" size="10"></TD>
    </TR>
<TR bgcolor="#E5E5E5"> 
      <TD><font color="darkblue" size="2" face="arial">Aadhaar No.(12 Digit)</font></TD>
      <TD><INPUT NAME="aadno" TYPE="text" maxlength="12" size="12"></TD>
    </TR>
    <TR bgcolor="#E5E5E5"> 
      <TD colspan=2 align=center><INPUT TYPE="submit" value=" Submit " onclick="return chk()"></TD>
    </TR>
  </TABLE>
  <blockquote>
    <p align="center"> <font color="#FF9999"><strong><font color="#CC0000">Mobile  
      No. and Date of Birth should be matched</font></strong></font></p>
  </blockquote>
</FORM>
</BODY>
</HTML>
