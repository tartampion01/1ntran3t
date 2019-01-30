<%Response.Expires=-1441%>
<%
'Any Users
PageLevel=0
%>
<!--#include virtual="/intranet/includes/check.asp" -->

<%

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application("cstring")

If request.form <> "" then

	ID = Request.Form("ID")
	Name = Request.Form("Name")
	category_id = Request.form("c")
	
		If category_id <> 0 then
			fldname = "product_id"
		Else
			fldname = "acc_id"
		End If	
	
	If request.Form("update")<>"" then
		Conn.execute("Delete from products_acc WHERE " & fldname & "='" & ID & "';")
		For each pid in request.form("pid")

		If category_id <> 0 then
			strSQL = strSQL & "INSERT INTO products_acc VALUES(" & ID & "," & pid & ");" & vbCrLf
		Else
			strSQL = strSQL & "INSERT INTO products_acc VALUES(" & pid & "," & ID & ");" & vbCrLf
		End If				
				
				Conn.Execute(strSQL)
				strSQL = ""
		Next
	'response.end
		If strSQL <> "" then Conn.Execute(strSQL)
		response.Redirect("attach.asp?id=" & ID & "&c=" & category_id & "&name=" & name)
	End if	
	
Else
	ID = Request.QueryString("ID")
	Name = Request.QueryString("Name")
	category_id = Request.QueryString("c")
End If

sql = "SELECT products.id,CONCAT(brandname,' ',Model,' ',Color) as Model,product_id,acc_id from products inner join brands ON brand_id=brands.id left join products_acc "
If category_id <> 0 then
	sql = sql & " on products.id = acc_id AND product_id='" & ID & "' WHERE category_id = 0"
	fldname = "acc_id"
Else
	sql = sql & " on products.id = product_id AND acc_id='" & ID & "'  WHERE category_id <> 0"
	fldname = "product_id"
End If
sql = sql & " ORDER By Model;"

set rs = conn.execute(sql)

%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/backend.dwt.asp" codeOutsideHTMLIsLocked="false" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Intranet - <%=dealers(session("dealer_id"),1)%></title>
<!-- InstanceEndEditable --><!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
<link href="/intranet/includes/styles_app.css" rel="stylesheet" type="text/css" />
<!-- InstanceParam name="id" type="text" value="products" -->
</head>
<body id="products">
<div id="container">
<div id="header">
  <ul id="nav">
    <li id="invent_tab"><a href="/intranet/inventory/">Inventaire</a></li>
	<%IF Session("Level")=1 then%>
	<li id="users_tab"><a href="/intranet/users/">Utilisateurs</a></li>
	<%End If%>
  </ul>
<ul id="second_nav"><li><em><%="<strong>" & Session("FullName") & "</strong> (" & UserLevel(session("level")) & ")"%></em></li><li><a href="/intranet/login/password.asp">Mot de passe</a></li>
<li><a href="/intranet/login/logout.asp">Quitter</a></li>
</ul>
</div><div id="content">
<!-- InstanceBeginEditable name="content" -->
 <h1>Attach to: <%=name%> </h1>
 <table border="0" width="100%" cellspacing="0" cellpadding="3" >
   <form method="post" action="attach.asp?id=">
    <input type="hidden" name="ID" value="<%=ID%>" />
	<input type="hidden" name="c" value="<%=category_id%>" />
	<input type="hidden" name="name" value="<%=name%>" />
 <tr> 
    <th valign="top" align="left" nowrap="nowrap" colspan="3"><a href="details.asp?id=<%=ID%>"><img src="/intranet/images/ico_fup.gif" width="16" height="16" border="0" /></a></th>
  </tr>
  <tr class="titlebar">
    <td align="left" nowrap="nowrap">Products</td>
    <td align="center" nowrap="nowrap">V</td>
    <td width="100%" align="left" nowrap="nowrap">&nbsp;</td>
  </tr>
<%

	Total = 0
  	Do While Not rs.EOF
	
			ct = ct + 1
			If ct/2 = Cint(ct/2) then
				TRStyle = "normal"
			Else
				TRStyle = "high"
			End If
	
			Total = total + 1
  
%>
  <tr class="<%=TRStyle%>">
    <td align="left" nowrap="nowrap"><%=rs("model")%></td>
    <td align="left" nowrap="nowrap"><input type="checkbox" name="pid" value="<%=rs("id")%>" <%if rs(fldname) > 0 then response.write "checked"%> /></td>
    <td align="left" nowrap="nowrap">&nbsp;</td>
  </tr>
<%
  			
    	rs.moveNext	
  	Loop

Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>
    <tr class="titlebar">
      <td align="left">Action</td>
      <td colspan="2" align="left">        <input type="submit" value="Update" name="update" class="button" />      </td>
      </tr>
  </form>
</table> 
<!-- InstanceEndEditable --></div></div>
</body>
<!-- InstanceEnd --></html>
