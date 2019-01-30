<%@EnableSessionState=False%>
<html>

<head>

<title>Uploading...</title>
<meta http-equiv="expires" content="Tue, 01 Jan 1981 01:00:00 GMT">
<meta http-equiv=refresh content="2,progressbar.asp?ID=<%=Request.QueryString("ID")%>">

<%
On Error Resume Next
Set theProgress = Server.CreateObject("ABCUpload4.XProgress")
theProgress.ID = Request.QueryString("ID")
%>

<script language="javascript">
<!--
if (<% =theProgress.PercentDone %> == 100) top.close();
if ("<% =theProgress.Note %>" == "Upload stopped unexpectedly. ") top.close();
//-->
</script>


<link rel="stylesheet" href="/intranet/includes/styles_app.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <th colspan="2"><img src="/intranet/images/clock.gif" width="11" height="11" align="absmiddle"> Uploading...</th>
  </tr>
  <tr>
    <td colspan="2">
      <table border="0" width="<%=theProgress.PercentDone%>%" cellspacing="0">
        <tr>
          <td class="cell_label">&nbsp;</td>
        </tr>
      </table>    </td>
  </tr>
  <tr>
    <td class="cell_label" nowrap>Time remaining:</td>
    <td width="100%">&nbsp;<%=Int(theProgress.SecondsLeft / 60)%> min<%=theProgress.SecondsLeft Mod 60%> sec 
      (<%=Round(theProgress.BytesDone / 1024, 1)%> KB of <%=Round(theProgress.BytesTotal / 1024, 1)%> KB transfered)</td>
  </tr>
  <tr>
    <td class="cell_label" nowrap> Transfer rate:</td>
    <td>&nbsp;<%=Round(theProgress.BytesPerSecond/1024, 1)%> KB/sec</td>
  </tr>
  <tr>
    <td class="cell_label" nowrap>Information:</td>
    <td>&nbsp;<%=theProgress.Note%></td>
  </tr>
  <tr>
    <td class="cell_label" nowrap>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  
  <tr></tr>
</table> 
</body>

</html>
