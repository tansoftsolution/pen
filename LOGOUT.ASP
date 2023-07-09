<%response.buffer=true %>
<% on error resume next%>
<html>
<BODY BGCOLOR="#00CCFF" text="#330066" link="#990000">
<%
session.abandon
%>

<div align="center">
  <h1>&nbsp;JPLC - AADHAAR COUNT SCREEN</h1>
</div>
<TR vAlign = middle height = 100><!-- row 3 -->
  <TD align=center>
    <div align="center">
      <p><FONT SIZE="2" COLOR="red" face="arial"><B><font size="3">You Have Successfully
        Logged Out </font></B></FONT></p>
      <p>&nbsp;</p>
    </div>
  </TD>
<TR>
  <TD align=center>
    <div align="center"><A HREF="/PEN/AADCOUNT.html"><font size="4">Login</font></A></div>
  </TD>
</TR>
</BODY>
</HTML>
