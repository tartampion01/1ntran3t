<%Response.Expires=-1441%>
<%
'Admin Only
PageLevel=1
%>
<!--#include virtual="/intranet/includes/check.asp" -->
<%
tableName = "intranet_users"
Title = "Liste des utilisateurs"
SearchItems=""
IF session("dealer_id")<>0 then DefaultWHERE = "dealer_id=" & session("dealer_id")

dim Headers(4,5)
Headers(0,0)="LastName"
Headers(0,1)="Nom"
Headers(0,2)=false
Headers(0,3)="rs(""FirstName"") & ""&nbsp;"" & rs(""LastName"")"

Headers(1,0)="phone"
Headers(1,1)="Téléphone"
Headers(1,2)=false

Headers(2,0)="email"
Headers(2,1)="Courriel"
Headers(2,2)=false

Headers(3,0)="dealer_id"
Headers(3,1)="Succursale"
Headers(3,2)=false
Headers(3,3)="Dealers(rs(""Dealer_id""),0)"

Headers(4,0)="Level"
Headers(4,1)="Type d'utilisateur"
Headers(4,2)=false
Headers(4,3)="UserLevel(rs(""Level""))"

strSQL = "Select * from intranet_users "

overdue_condition = "false"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/backend.dwt.asp" codeOutsideHTMLIsLocked="false" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Intranet - <%=dealers(session("dealer_id"),1)%></title>
<!-- InstanceEndEditable --><!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
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
	</ul>
	<ul id="second_nav">
		<li><em><%="<strong>" & Session("FullName") & "</strong> (" & UserLevel(session("level")) & ")"%></em></li>
		<li><a href="/intranet/login/password.asp">Mot de passe</a></li>
		<li><a href="/intranet/login/logout.asp">Quitter</a></li>
	</ul>
</div>
<div id="content">
<!-- InstanceBeginEditable name="content" -->
 
<!--#include virtual="/intranet/includes/list_inc.asp" -->
 
<!-- InstanceEndEditable --></div></div>
</body>
<!-- InstanceEnd --></html>

