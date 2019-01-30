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
tableName = "inventory"
Title = "Inventaires"
SearchItems="model,stock,color"
uniqueListName = "inventory"
DefaultOrderBy = "dealer_id"

'Fields to display

dim Headers()

if isAdmin then 
	redim Headers(17,5)
else
	redim Headers(16,5)
	DefaultWHERE = "ordersold=''"
end if

Headers(0,0)="stock"
Headers(0,1)="Stock"
Headers(0,2)=false

Headers(1,0)="serial"
Headers(1,1)="#Série"
Headers(1,2)=false

i=0
if isAdmin then 
	Headers(2,0)="ordersold"
	Headers(2,1)="Comm Vendu"
	Headers(2,2)=false	
	i=1
end if

Headers(i+2,0)="ordernumber"
Headers(i+2,1)="Comm Stock"
Headers(i+2,2)=false

Headers(i+3,0)="marque"
Headers(i+3,1)="Marque"
Headers(i+3,2)=true
Headers(i+3,4)="Select DISTINCT marque from inventory order by marque;"
Headers(i+3,5)="rsCombo(""marque"")"

Headers(i+4,0)="model"
Headers(i+4,1)="Modèle"
Headers(i+4,2)=true
Headers(i+4,4)="Select DISTINCT model from inventory order by model;"
Headers(i+4,5)="rsCombo(""model"")"

Headers(i+5,0)="dealer_id"
Headers(i+5,1)="Succ."
Headers(i+5,2)=true
Headers(i+5,3)="dealers(rs(""dealer_id""),0)"
Headers(i+5,4)="Select DISTINCT dealer_id from inventory order by dealer_id;"
Headers(i+5,5)="dealers(rsCombo(""dealer_id""),0)"

Headers(i+6,0)="engine"
Headers(i+6,1)="Moteur"
Headers(i+6,2)=true
Headers(i+6,4)="Select DISTINCT engine from inventory order by engine;"
Headers(i+6,5)="rsCombo(""engine"")"


dim check(1)
	check(0)=""
	check(1)="<input type=checkbox checked />"

Headers(i+7,0)="transmission"
Headers(i+7,1)="Transmission"
Headers(i+7,2)=false

Headers(i+8,0)="axle"
Headers(i+8,1)="Essieux"
Headers(i+8,2)=false

Headers(i+9,0)="rearsuspension"
Headers(i+9,1)="Susp."
Headers(i+9,2)=false

Headers(i+10,0)="wb"
Headers(i+10,1)="WB"
Headers(i+10,2)=false

Headers(i+11,0)="ratio"
Headers(i+11,1)="Ratio"
Headers(i+11,2)=false

Headers(i+12,0)="color"
Headers(i+12,1)="Couleur"
Headers(i+12,2)=false

Headers(i+13,0)="demo"
Headers(i+13,1)="Démo"
Headers(i+13,2)=false

Headers(i+14,0)="nointerestDate"
Headers(i+14,1)="Date de prod."
Headers(i+14,2)=false
Headers(i+14,3)="makelocalDate(rs(""nointerestDate""))"

Headers(i+15,0)="location"
Headers(i+15,1)="Lieu"
Headers(i+15,2)=true
Headers(i+15,4)="Select DISTINCT location from inventory order by location;"
Headers(i+15,5)="rsCombo(""location"")"

Headers(i+16,0)="daysin"
Headers(i+16,1)="# jours"
Headers(i+16,2)=false
Headers(i+16,3)="ShowDays(rs(""daysin""))"

overdue_condition = "rs(""resuserid"")"

strSQL = "SELECT *,CONCAT(frontaxle,'/',rearaxle) as axle, CONCAT(transtype,' ',transoption) as transmission,IF(bonus>0,CONCAT('<span class=err>',cast(bonus as CHAR),'$</span>'),'-') AS bonus, cast(ifnull(DATEDIFF(CURDATE(),invoicedate),0) as SIGNED) as daysin from inventory "


function ShowDays(d)
	ShowDays = "N/D"
	if cint(d)>0 then
		ShowDays = "<span class='err'>" & d & "</span>"
	end if
end function


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/backend.dwt.asp" codeOutsideHTMLIsLocked="false" -->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<!-- InstanceBeginEditable name="doctitle" -->
<title>Intranet - <%=dealers(session("dealer_id"),1)%></title>
<!-- InstanceEndEditable --><!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable -->
<link href="/intranet/includes/styles_app.css" rel="stylesheet" type="text/css" />
<!-- InstanceParam name="id" type="text" value="inventory" -->
</head>
<body id="inventory">
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
<!--#include virtual="/intranet/includes/list_inc.asp" -->


<%
response.write err.description
%>

<!-- InstanceEndEditable --></div></div>
</body>
<!-- InstanceEnd --></html>
