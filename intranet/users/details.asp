<%Response.Expires=-1441%>
<%
'Admin Only
PageLevel=1
%>
<!--#include virtual="/intranet/includes/check.asp" -->
<!-- #include virtual="/intranet/includes/functions.asp" -->
<%
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application("cstring")

ID = Request.QueryString("ID")
IF ID="" then ID=Request.Form("ID")

' New User
If Request.Form<> "" And ID="" then
	on error resume next
	Conn.Execute("INSERT intranet_users (username) VALUES ('" & request.form("username") & "');")
	Set rs = Conn.Execute("SELECT LAST_INSERT_ID() AS ID;")
	ID = rs("ID")
	Set rs=Nothing
	if err <> 0 then 
		id=""
		errMsg = errMsg & "<span class=err>Le nom d'usager est déjà utilisé.</span>"
	End If
End IF

' Delete User
IF Request.QueryString("action")="del" then
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Application("cstring")
	
	Conn.Execute("Delete from intranet_users Where ID=" & ID & ";")
		
	Conn.Close
	Set Conn=Nothing
	
	Response.Redirect("default.asp")
End IF

' Modify User
If Request.Form("modUser")<>"" AND ID <> "" then

	strSQL = "Update intranet_users SET" & _
	" FirstName='" & Replace(Request.Form("FirstName"),"'","''") & "'," &_
	" LastName='" & Replace(Request.Form("LastName"),"'","''") & "'," &_
	" phone='" & Replace(Request.Form("phone"),"'","''") & "'," &_
	" Level=" & Request.Form("level") & "," &_
	" dealer_id=" & Request.Form("dealer_id") & "," &_
	" email='" & Replace(Request.Form("email"),"'","''") & "'"

	If Request.Form("password")<> "" then
		If 	Request.Form("password")=Request.Form("password1") then
			strSQL = strSQL & ", password='" & encode(Replace(Request.Form("password"),"'","''")) & "'"
			errMsg = errMsg & "<span class=ok>La mot de passe a été modifié avec succès.</span>"
		Else
			errMsg = errMsg & "<span class=err>La vérification du mot de passe a échouée.</span>"
		End If
	End If

	strSQL = strSQL & " WHERE ID=" & ID & ";"

	Conn.Execute(strSQL)
End If

HaveSaveRights = True

If ID <> "" then
	strSQL = "Select * from intranet_users WHERE ID=" & ID & ";"
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open strSQL,Conn
		
	For each item in rs.Fields
		Execute(item.Name & "=rs(""" & item.Name & """)")	
	Next
	rs.Close
	
	if dealer_id <> session("dealer_id") AND session("dealer_id")<>0 then HaveSaveRights = False
	
End If

Set rs=Nothing
Conn.Close
Set Conn=Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/backend.dwt.asp" codeOutsideHTMLIsLocked="false" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Intranet - <%=dealers(session("dealer_id"),1)%></title>
<!-- InstanceEndEditable --><!-- InstanceBeginEditable name="head" -->
 

<script language="JavaScript" type="text/javascript">
<!--
function Delete(ID) {
	if (confirm("ATTENTION!\nCette action supprimera définitivement cette fiche.\nVoulez-vous continuer?")){
		string="details.asp?action=del&ID="+ID;
		document.location.href=string;
	}
}
// -->
</script>

<link rel="stylesheet" href="/intranet/includes/styles_app.css" type="text/css" />
<!-- InstanceEndEditable -->
<link href="/intranet/includes/styles_app.css" rel="stylesheet" type="text/css" />
<!-- InstanceParam name="id" type="text" value="users" -->
</head>
<body id="users">
<div id="container">
<div id="header">
	<ul id="nav">
		<!-- UTILISATEUR -->
		<%IF Session("Level")=0 then%>
		<li id="invent_tab"><a href="/intranet/inventory/">Inventaire</a></li>
		<%End If%>
		<!-- ADMINISTRATEUR -->
		<%IF Session("Level")=1 then%>
		<li id="invent_tab"><a href="/intranet/inventory/">Inventaire</a></li>
		<li id="users_tab"><a href="/intranet/users/">Utilisateurs</a></li>
		<li id="emplois_tab"><a href="/intranet/emplois/">Emplois</a></li>
		<%End If%>
		<!-- RH -->
		<%IF Session("Level")=2 then%>
		<li id="emplois_tab"><a href="/intranet/emplois/">Emplois</a></li>
		<%End If%>
		<!-- ReadOnly -->
		<%IF Session("Level")=3 then%>
		<li id="invent_tab"><a href="/intranet/inventory/">Inventaire</a></li>
		<%End If%>
	</ul>
	<ul id="second_nav">
		<li><em><%="<strong>" & Session("FullName") & "</strong> (" & UserLevel(session("level")) & ")"%></em></li>
		<li><a href="/intranet/login/password.asp">Mot de passe</a></li>
		<li><a href="/intranet/login/logout.asp">Quitter</a></li>
	</ul>
</div>
<div id="content">
<!-- InstanceBeginEditable name="content" -->
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top"><h1>Utilisateur: <%=username%></h1>
        <%If errMsg <> "" then%>
      <%=ErrMsg%>
        <%End IF%></td>
  </tr>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="3" >
  <tr> 
    <th valign="top" align="left" nowrap="nowrap" colspan="2"><a href="default.asp"><img src="/intranet/images/ico_fup.gif" width="16" height="16" border="0" /></a></th>
  </tr>
  <tr class="titlebar">
    <td align="left" nowrap="nowrap" colspan="2">Identification</td>
  </tr>
  <form method="post" action="/intranet/users/details.asp">
    <input type="hidden" name="ID" value="<%=ID%>" />
    <tr>
      <td height="100%" class="cell_label">Succursale :</td>
      <td class="cell_content"><select name="dealer_id">
          <%=DealerCombo(dealer_id)%>
      </select></td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Nom :</td>
      <td width="100%" class="cell_content"><input type="text" name="LastName" size="40" value="<%=LastName%>" /></td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Pr&eacute;nom :</td>
      <td width="100%" class="cell_content"> 
        <input type="text" name="FirstName" size="40" value="<%=FirstName%>" />      </td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Type:</td>
      <td width="100%" class="cell_content"> 
        <select name="level">			
          <option <%If Level=0 then response.write "selected"%> value="0">Utilisateur</option>			
          <option <%If Level=1 then response.write "selected"%> value="1">Administrateur</option>
		  <option <%If Level=2 then response.write "selected"%> value="2">RH</option>
		  <option <%If Level=3 then response.write "selected"%> value="3">Lecture Seule</option>
        </select>
	</td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Nom d'utilisateur :</td>
      <td width="100%" class="cell_content"> 
        <input type="text" name="username" size="40" value="<%=username%>" <%if ID <> "" then response.write "readonly" %> />      </td>
    </tr>
    <tr> 
      <td valign="top" align="left" nowrap="nowrap" class="cell_label">Mot de passe :</td>

	  
      <td width="100%" class="cell_content"> 
       <input name="password" type="password" value="" size="15" />
        <br />
        <input name="password1" type="password" value="" size="15" />
(v&eacute;rification)</td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">T&eacute;l&eacute;phone:</td>
      <td width="100%" class="cell_content"> 
        <input type="text" name="Phone" size="40" value="<%=Phone%>" />      </td>
    </tr>
    <tr> 
      <td align="left" nowrap="nowrap" class="cell_label">Courriels:</td>
      <td width="100%" class="cell_content"> 
        <input type="text" name="Email" size="40" value="<%=Email%>" />
      <a href="mailto:<%=Email1%>"><img src="/intranet/images/ico_email.gif" width="23" height="22" align="absmiddle" border="0" alt="Send an email" /></a>	  </td>
    </tr>
    <tr class="titlebar">
      <td align="left" colspan="2">Actions</td>
    </tr>

<%
If HaveSaveRights then
	buttonstate = ""
Else
	buttonstate = "disabled"
End If

%>

    <tr> 
      <td valign="top" align="left" nowrap="nowrap" class="cell_label">&nbsp;</td>
      <td valign="top" width="100%" align="left" class="cell_content"> 
        <input type="submit" value="Sauvegarder" name="modUser" class="button" <%=buttonstate%> />
        <% If ID <> "" then%>
        <input onclick="Delete(<%=ID%>)" type="button" value="Supprimer" name="del" class="button" <%=buttonstate%> />
        <%End If%>      </td>
    </tr>
  </form>  
</table> 
<span style="color: #FFFFFF;visibility:hidden">
	<% Response.Write(Decode(password))%>
</span>
<!-- InstanceEndEditable --></div></div>
</body>
<!-- InstanceEnd --></html>