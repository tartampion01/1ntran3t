<%Response.Expires=-1441%>
<%
'Any Users
PageLevel=0

isAdmin = Session("Level")

on error resume next

%>
<!-- #include virtual="/intranet/includes/check.asp" -->
<%

'Table to display
tableName = "offresEmplois"
Title = "Emplois"
SearchItems="titre,succursale,contactNom"
uniqueListName = "offresEmplois"
DefaultOrderBy = "emploi_id"

'Fields to display

dim Headers()

if isAdmin then 
	redim Headers(7,5)
else
	redim Headers(7,5)
	'DefaultWHERE = "1=1"
end if

Headers(0,0)="emploi_id"
Headers(0,1)="Id"
Headers(0,2)=false

Headers(1,0)="titre"
Headers(1,1)="Titre"
'Headers(1,2)=false
Headers(1,2)=true
Headers(1,4)="Select DISTINCT titre from offresEmplois order by titre;"
Headers(1,5)="rsCombo(""titre"")"

Headers(2,0)="fonctions"
Headers(2,1)="Fonctions"
Headers(2,2)=false

Headers(3,0)="referenceInterne"
Headers(3,1)="Référence Interne"
Headers(3,2)=false

Headers(4,0)="contactNom"
Headers(4,1)="Nom contact"
'Headers(4,2)=false
Headers(4,2)=true
Headers(4,4)="Select DISTINCT contactNom from offresEmplois order by contactNom;"
Headers(4,5)="rsCombo(""contactNom"")"

Headers(5,0)="succursale"
Headers(5,1)="Succursale"
'Headers(5,2)=false
Headers(5,2)=true
Headers(5,4)="Select DISTINCT succursale from offresEmplois order by succursale;"
Headers(5,5)="rsCombo(""succursale"")"

Headers(6,0)="dateDebut"
Headers(6,1)="Date début"
Headers(6,2)=false

Headers(7,0)="filled"
Headers(7,1)="Comblée"
Headers(7,2)=false

strSQL = "SELECT * from offresEmplois "

function ShowDays(d)
	ShowDays = "N/D"
	if cint(d)>0 then
		ShowDays = "<span class='err'>" & d & "</span>"
	end if
end function

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/backend.dwt.asp" codeOutsideHTMLIsLocked="false" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Intranet - <%=dealers(session("dealer_id"),1)%></title>
<!-- InstanceEndEditable --><!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
<link href="/intranet/includes/styles_app.css" rel="stylesheet" type="text/css" />
<!-- InstanceParam name="id" type="text" value="inventory" -->
</head>
<body id="emplois">
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
<!--#include virtual="/intranet/includes/list_emp.asp" -->

<%
response.write err.no & "   " & err.description
%>

<!-- InstanceEndEditable --></div></div>
</body>
<!-- InstanceEnd --></html>
